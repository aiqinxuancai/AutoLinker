// Run with the path to an installed playwright or playwright-core package.
const assert = require('node:assert/strict');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { chromium } = require(process.argv[2] || 'playwright');

(async () => {
	const browser = await chromium.launch({ headless: true });
	try {
		for (const file of ['ai_chat_history.html', 'ai_chat_history.src.html']) {
			for (const viewport of [{ width: 1000, height: 800 }, { width: 390, height: 700 }]) {
				const page = await browser.newPage({ viewport });
				const errors = [];
				page.on('pageerror', error => errors.push(error.message));
				await page.goto(pathToFileURL(path.resolve(__dirname, '../src/webview', file)).href);
				const result = await page.evaluate(async () => {
					const lines = Array.from({ length: 100 }, (_, i) => 'Output line ' + i).join('\n');
					const message = content => '<section class="msg tool"><div class="body">' + content + '</div></section>';
					const html = tick => message('<div class="plan-card tool-transcript-card"><pre class="tool-transcript-io">' + lines + '</pre></div>') +
						message('<div class="code-block"><pre id="stream-code" style="max-height:100px">' + lines + '\n' + tick + '</pre></div>') +
						message('<div class="markdown-table-wrap"><table style="width:1800px"><tr><td>Wide table</td></tr></table></div>') +
						'<section class="msg assistant plan-message"><div class="body"><div class="plan-card"><div class="plan-content">' +
						Array.from({ length: 30 }, (_, i) => '<div class="plan-step pending">Step ' + i + '</div>').join('') +
						'<p>' + tick + '</p></div></div></div></section>';
					window.autolinkerPlanDockExpanded = true;
					window.autolinkerSetChatHtml(html(0), 'session-a');
					const pending = Array.from({ length: 20 }, (_, i) => ({ text: 'Pending ' + i }));
					window.autolinkerSetPendingInputs(pending);
					const selectors = ['.tool-transcript-io', '#stream-code', '.markdown-table-wrap', '.activity-plan-details', '.pending-inputs-list'];
					const nodes = selectors.map(selector => document.querySelector(selector));
					nodes.forEach(node => { node.scrollTop = 45; node.scrollLeft = 80; });
					const positions = nodes.map(node => [node.scrollTop, node.scrollLeft]);
					if (positions.some(([top, left]) => top === 0 && left === 0)) throw new Error('Fixture is not scrollable');
					const history = document.getElementById('history-scroll');
					history.scrollTop = 0;
					for (let tick = 1; tick <= 25; tick++) {
						window.autolinkerSetChatHtml(html(tick), 'session-a');
						window.autolinkerSetChatHtml(html(tick), 'session-a');
						window.autolinkerSetPendingInputs(pending.concat({ text: 'Next ' + tick }));
						await new Promise(requestAnimationFrame);
						selectors.forEach((selector, i) => {
							const node = document.querySelector(selector);
							if (node !== nodes[i] || node.scrollTop !== positions[i][0] || node.scrollLeft !== positions[i][1]) {
								throw new Error('Scroll or node reset: ' + selector + ' at update ' + tick);
							}
						});
						if (history.scrollTop !== 0) throw new Error('History jumped while reading');
					}
					if (!document.querySelector('#stream-code').textContent.endsWith('25')) throw new Error('Stream did not update');
					history.scrollTop = history.scrollHeight;
					window.autolinkerSetChatHtml(html(26) + message('<p>New message</p>'), 'session-a');
					if (history.scrollHeight - history.scrollTop - history.clientHeight > 1) throw new Error('Bottom following failed');
					window.autolinkerSetChatHtml(html(26), 'session-b');
					if (document.querySelector('.tool-transcript-io') === nodes[0] || document.querySelector('.tool-transcript-io').scrollTop !== 0) {
						throw new Error('Session switch retained previous state');
					}
					window.autolinkerSetChatHtml('', 'session-b');
					if (document.getElementById('chat-root').childNodes.length || !document.getElementById('active-plan-panel').hidden) throw new Error('Clear failed');
					return { updates: 50, positions };
				});
				assert.deepEqual(errors, []);
				console.log(file, viewport, result);
				await page.close();
			}
		}
	} finally {
		await browser.close();
	}
})().catch(error => { console.error(error); process.exitCode = 1; });
