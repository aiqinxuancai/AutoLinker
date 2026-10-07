"""在已打开的 e-window-exe-full.e 临时副本上验证未保存代码和事件的完整镜像。"""

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import struct
import xml.etree.ElementTree as ET

from test_const_page_write import Mcp, source


def event_indices(path):
    data = path.read_bytes()
    assert data[:8] == b'CNWTEPRG'
    offset = 8
    while offset < len(data):
        assert struct.unpack_from('<I', data, offset)[0] == 353465113
        key = struct.unpack_from('<I', data, offset + 8)[0]
        size = struct.unpack_from('<i', data, offset + 56)[0]
        if key == 0x0A007319:
            return list(struct.iter_unpack('<iiii', data[offset + 100:offset + 100 + size]))
        offset += 100 + size
    return []


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--project', type=Path, required=True, help='可丢弃的测试副本，禁止使用原工程')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--endpoint', default='http://127.0.0.1:19207/mcp')
    args = parser.parse_args()
    project = args.project.resolve()
    args.output.mkdir(parents=True, exist_ok=True)
    disk_hash = hashlib.sha256(project.read_bytes()).hexdigest()
    expected_count = len(event_indices(project))
    assert expected_count > 0, 'Test project must have indexed events'
    mcp = Mcp(args.endpoint, args.output)
    mcp.rpc('initialize', {'protocolVersion': '2024-11-05', 'capabilities': {},
                          'clientInfo': {'name': 'window-event-mirror-regression', 'version': '1'}})
    matches = [i for i in mcp.call('list_instances')['instances']
               if Path(i['source_file_path']).resolve() == project]
    assert len(matches) == 1, matches
    instance = matches[0]
    assert mcp.call('select_instance', instance_id=instance['instance_id'])['ok']
    code_path = 'src/窗口程序集_启动窗口.txt'
    results = []

    def capture(label, mode, marker):
        refresh = mcp.call('refresh_workspace_mirror', mode=mode)
        assert refresh['ok'], refresh
        assert Path(refresh['source_file_path']).resolve() == project, refresh
        mirror = Path(refresh['mirror_root'])
        destination = args.output / label
        shutil.copytree(mirror, destination)
        native = destination / 'project/.native_source.bin'
        indexed = event_indices(native)
        methods = {method['id'] for page in json.loads(
            (destination / 'project/.native_source_map.json').read_text(encoding='utf-8-sig'))
            for method in page['methods']}
        assert len(indexed) == expected_count, (label, indexed)
        assert len({entry[:3] for entry in indexed}) == expected_count, indexed
        assert all(entry[3] in methods for entry in indexed), indexed
        events = [node for file in (destination / 'src').rglob('*.xml')
                  for node in ET.parse(file).iter() if node.tag.endswith('.事件')]
        assert len(events) == expected_count, (label, len(events))
        assert all('::' in node.get('处理器', '') for node in events)
        assert marker in (destination / code_path).read_text(encoding='utf-8-sig')
        assert hashlib.sha256(project.read_bytes()).hexdigest() == disk_hash, 'Refresh saved the source file'
        current = next(i for i in mcp.call('list_instances')['instances']
                       if i['instance_id'] == instance['instance_id'])
        assert Path(current['source_file_path']).resolve() == project, current
        results.append({'case': label, 'events': len(events), 'valid_indexed_handlers': len(indexed),
                        'source_path_preserved': True, 'disk_unchanged': True})
        print(label + ': ok', flush=True)

    capture('before-write', 'full', 'HAHAHA')
    baseline = mcp.call('read_real_file', file_path=code_path)
    original = source(baseline)
    assert 'HAHAHA' in original
    marker = 'ISSUE31_UNSAVED_MIRROR'
    result = mcp.call('write_file', file_path=code_path,
                      full_code=original.replace('HAHAHA', marker, 1),
                      expected_base_hash=baseline['code_hash'])
    assert result['ok'] and result['verified'], result
    capture('unsaved-main-only', 'main_only', marker)
    capture('unsaved-full', 'full', marker)
    (args.output / 'summary.json').write_text(json.dumps(
        {'ok': True, 'project': str(project), 'source_sha256': disk_hash, 'cases': results},
        ensure_ascii=False, indent=2), encoding='utf-8')
    print('Unsaved test edit remains in the disposable IDE project; source file was not saved.')


if __name__ == '__main__':
    main()
