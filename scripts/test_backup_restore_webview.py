"""使用本机 Edge 验证备份恢复页面；需要 playwright（可安装到 .tmp/browser-test-deps）。"""
import sys,json
from pathlib import Path
sys.path.insert(0,str(Path('.tmp/browser-test-deps').resolve()))
from playwright.sync_api import sync_playwright
out=Path('docs/test-results/backup-restore-webview');out.mkdir(parents=True,exist_ok=True)
with sync_playwright() as p:
 b=p.chromium.launch(channel='msedge',headless=True)
 page=b.new_page(viewport={'width':940,'height':680})
 page.add_init_script("window.messages=[];window.chrome={webview:{postMessage:m=>window.messages.push(JSON.parse(m))}}")
 page.goto(Path('src/webview/backup_restore_settings.html').resolve().as_uri())
 payload={'supported':True,'canRestore':True,'revision':42,'sourcePath':r'D:\工程\测试.e','backups':[{'name':'20261009120000-[修改中文😀&需求摘要].e','timestamp':'2026-10-09 12:00:00','size':138133}]}
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 assert page.locator('#rows td').nth(1).inner_text()==payload['backups'][0]['name']
 page.screenshot(path=str(out/'page.png'))
 page.on('dialog',lambda d:d.accept())
 page.get_by_role('button',name='恢复',exact=True).click()
 assert page.evaluate('window.messages.at(-1)')=={'action':'restore','index':0,'revision':42}
 assert page.locator('#refresh').is_disabled()
 payload['notice']='恢复成功。\n恢复前安全备份：测试.e'
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 assert page.locator('#refresh').is_enabled()
 payload['supported']=False
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 assert page.get_by_role('button',name='恢复',exact=True).is_disabled()
 assert 'e5.95' in page.locator('#support').inner_text()
 payload['supported']=True;payload['canRestore']=False
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 assert page.get_by_role('button',name='恢复',exact=True).is_disabled()
 payload['backups']=[]
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 assert '暂无' in page.locator('#rows').inner_text()
 payload['canRestore']=True
 payload['backups']=[{'name':'chat.e','kind':0},{'name':'timer.e','kind':1},{'name':'tool.e','kind':2}]
 payload['settings']=[{'enabled':True,'keepCount':10},{'enabled':False,'keepCount':20},{'enabled':False,'keepCount':20}]
 payload['intervalMinutes']=5
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 page.get_by_role('tab',name='定时快照',exact=True).click()
 assert page.locator('#rows td').nth(1).inner_text()=='timer.e'
 assert page.locator('#interval').is_visible()
 page.locator('#enabled').check();page.locator('#keep').fill('3');page.locator('#interval').fill('1')
 page.get_by_role('button',name='保存设置',exact=True).click()
 assert page.evaluate('window.messages.at(-1)')=={'action':'save','kind':1,'enabled':True,'keepCount':3,'intervalMinutes':1}
 n=page.evaluate('window.messages.length');page.locator('#keep').fill('0')
 page.get_by_role('button',name='保存设置',exact=True).click()
 assert page.evaluate('window.messages.length')==n
 page.get_by_role('tab',name='MCP / 内部工具写入',exact=True).click()
 assert page.locator('#rows td').nth(1).inner_text()=='tool.e'
 assert page.locator('#interval').is_hidden()
 restore=page.get_by_role('button',name='恢复',exact=True);restore.hover()
 colors=restore.evaluate('(e)=>({fg:getComputedStyle(e).color,bg:getComputedStyle(e).backgroundColor})')
 assert colors['fg']!=colors['bg']
 restore.click()
 assert page.evaluate('window.messages.at(-1)')=={'action':'restore','index':2,'revision':42}
 assert page.get_by_role('tab',name='定时快照',exact=True).is_disabled()
 page.evaluate('d=>window.autolinkerApplyBackupRestore(d)',payload)
 assert page.get_by_role('tab',name='定时快照',exact=True).is_enabled()
 page.screenshot(path=str(out/'three-tabs.png'))
 b.close()
(out/'result.json').write_text(json.dumps({'passed':True,'browser':'Edge headless','checks':['Unicode rendering','restore index and revision','busy state','completion resets controls','unsupported IDE disables restore','non .e disables restore','empty list','category filters','timed settings','invalid retention rejected','original restore index after filtering','hover contrast','tabs disabled during restore']},ensure_ascii=False,indent=2),encoding='utf-8')
print('PASS: browser regression checks')
