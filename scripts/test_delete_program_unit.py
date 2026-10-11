"""在已打开的临时工程副本上验证原生程序集/类删除；不保存工程，禁止指定原工程。"""

import argparse
import hashlib
import json
from pathlib import Path

from test_const_page_write import Mcp, source


class Client(Mcp):
    def call(self, tool, /, **arguments):
        return self.rpc('tools/call', {'name': tool, 'arguments': arguments})['structuredContent']


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--endpoint', default='http://127.0.0.1:19207/mcp')
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    project = args.project.resolve()
    disk_hash = hashlib.sha256(project.read_bytes()).hexdigest()
    mcp = Client(args.endpoint, args.output)
    mcp.rpc('initialize', {'protocolVersion': '2024-11-05', 'capabilities': {},
                          'clientInfo': {'name': 'delete-program-unit-regression', 'version': '1'}})
    instances = mcp.call('list_instances')['instances']
    matches = [i for i in instances if i.get('reachable') and Path(i['source_file_path']).resolve() == project]
    assert len(matches) == 1, instances
    assert mcp.call('select_instance', instance_id=matches[0]['instance_id'])['ok']
    catalog = mcp.rpc('tools/list', {})['tools']
    assert any(t['name'] == 'delete_file' for t in catalog), 'New plugin not loaded'
    refreshed = mcp.call('refresh_workspace_mirror', mode='full')
    assert refreshed['ok'], refreshed
    mirror = Path(refreshed['mirror_root'])
    baseline = {p.relative_to(mirror).as_posix(): p.read_bytes() for p in (mirror / 'src').rglob('*') if p.is_file()}
    cases = []

    def record(name):
        cases.append(name)
        print(name + ': ok', flush=True)

    def remove(path, expected_hash):
        return mcp.call('delete_file', file_path=path, expected_base_hash=expected_hash)

    def check_survivors():
        refresh = mcp.call('refresh_workspace_mirror', mode='full')
        assert refresh['ok'], refresh
        root = Path(refresh['mirror_root'])
        for path, content in baseline.items():
            assert (root / path).read_bytes() == content, path
        return root

    for kind in ('assembly', 'class'):
        # 刻意使名称误导类型猜测，类型必须来自镜像声明而非名称或树图标。
        name = '类名称的普通程序集' if kind == 'assembly' else '普通名称的对象'
        # 覆盖带程序集变量、多个子程序的页面；类生命周期由新增工具保留。
        code = '\n'.join(['.版本 2', '', '.程序集 ' + name,
                          '.程序集变量 保留成员, 整数型', '',
                          '.子程序 回归函数甲, 整数型, 公开', '返回 (11)', '',
                          '.子程序 回归函数乙, 整数型, 公开', '返回 (22)', ''])
        added = mcp.call('add_new_file', file_type=kind, name=name, full_code=code)
        assert added['ok'] and added['verified'], added
        path = added['file_path']
        base = mcp.call('read_real_file', file_path=path)
        assert '回归函数甲' in source(base) and '回归函数乙' in source(base), base
        stale = remove(path, '0' * 64)
        assert not stale['ok'] and not stale['deleted'] and 'expected_base_hash mismatch' in stale['error'], stale
        assert mcp.call('read_real_file', file_path=path)['code_hash'] == base['code_hash']
        record(kind + '_stale_hash_preserves_page')
        edited = mcp.call('edit_file', file_path=path, old_text='返回 (11)', new_text='返回 (12)',
                          expected_base_hash=base['code_hash'])
        assert edited['ok'] and edited['verified'], edited
        stale = remove(path, base['code_hash'])
        assert not stale['ok'] and not stale['deleted'], stale
        base = mcp.call('read_real_file', file_path=path)
        assert '返回 (12)' in source(base), base
        record(kind + '_changed_source_requires_new_hash')
        # 元数据缺失不能被当成“没有窗口绑定”；拒绝后保持源码和磁盘元数据原样。
        current_mirror = mcp.call('refresh_workspace_mirror', mode='full')
        assert current_mirror['ok'], current_mirror
        meta_path = Path(current_mirror['mirror_root']) / 'project' / '_meta.json'
        metadata_bytes = meta_path.read_bytes()
        metadata = json.loads(metadata_bytes.decode('utf-8-sig'))
        assert 'windowBindings' in metadata
        del metadata['windowBindings']
        try:
            meta_path.write_text(json.dumps(metadata, ensure_ascii=False), encoding='utf-8')
            rejected = remove(path, base['code_hash'])
            assert not rejected['ok'] and not rejected['deleted'], rejected
            assert 'program_unit_metadata_invalid' in rejected['trace'], rejected
        finally:
            meta_path.write_bytes(metadata_bytes)
        assert mcp.call('read_real_file', file_path=path)['code_hash'] == base['code_hash']
        record(kind + '_missing_window_metadata_rejected')
        # 建立另一页，使删除目标不依赖当前页或程序树选择。
        sentinel = mcp.call('add_new_file', file_type='assembly', name='删除时的活动页')
        assert sentinel['ok'], sentinel
        deleted = remove(path, base['code_hash'])
        assert deleted['ok'] and deleted['deleted'] and deleted['verified'], deleted
        assert deleted['file_type'] == kind and deleted['workspace_mirror_refreshed'], deleted
        assert 'native_delete_selected_program_item|command=0x82020017' in deleted['trace'], deleted
        assert 'ide_entry_probe' not in deleted['trace'], deleted
        assert mcp.call('read_real_file', file_path=sentinel['file_path'])['ok']
        root = check_survivors()
        assert not (root / path).exists(), path
        failed_read = mcp.call('read_real_file', file_path=path)
        assert not failed_read['ok'], failed_read
        again = remove(path, base['code_hash'])
        assert not again['ok'] and not again['deleted'], again
        record(kind + '_delete_nonactive_page_and_repeat')
        recreated = mcp.call('add_new_file', file_type=kind, name=name)
        assert recreated['ok'] and recreated['file_path'] == path, recreated
        current = mcp.call('read_real_file', file_path=path)
        assert '回归函数甲' not in source(current), current
        deleted = remove(path, current['code_hash'])
        assert deleted['ok'] and deleted['workspace_mirror_refreshed'], deleted
        sentinel_base = mcp.call('read_real_file', file_path=sentinel['file_path'])
        assert remove(sentinel['file_path'], sentinel_base['code_hash'])['ok']
        record(kind + '_recreate_same_name_and_delete_active_page')

    for path in ('src/.常量.txt', 'src/不存在的程序集.txt', '../delete-test.e'):
        rejected = remove(path, '0' * 64)
        assert not rejected['ok'] and not rejected['deleted'], rejected
    record('fixed_table_missing_file_and_path_traversal_rejected')
    # 窗口工程额外验证窗口绑定程序集和 XML 不会被删除。
    for path in baseline:
        if path.endswith('.xml'):
            assert not remove(path, '0' * 64)['ok']
        if path.endswith('.txt') and (path.startswith('src/窗口程序集') or (path[:-4] + '.xml') in baseline):
            base = mcp.call('read_real_file', file_path=path)
            rejected = remove(path, base['code_hash'])
            assert not rejected['ok'] and 'window_bound_assembly_not_supported' in rejected['trace'], rejected
            assert mcp.call('read_real_file', file_path=path)['code_hash'] == base['code_hash']
            record('window_bound_assembly_preserved')
    root = check_survivors()
    assert {p.relative_to(root).as_posix() for p in (root / 'src').rglob('*') if p.is_file()} == set(baseline)
    assert hashlib.sha256(project.read_bytes()).hexdigest() == disk_hash, 'Tool saved the source file'
    record('all_original_sources_preserved_and_disk_unchanged')
    executable = args.output.resolve() / 'after-delete.exe'
    compiled = mcp.call('compile_with_output_path', output_path=str(executable),
                        compiler='eide', static_compile=False)
    assert compiled['ok'] and executable.is_file(), compiled
    record('project_compiles_after_deletions')
    (args.output / 'summary.json').write_text(json.dumps(
        {'ok': True, 'project': str(project), 'source_sha256': disk_hash, 'cases': cases},
        ensure_ascii=False, indent=2), encoding='utf-8')


if __name__ == '__main__':
    main()
