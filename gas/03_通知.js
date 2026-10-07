/**
 * 出荷前の未チェック通知（時間主導トリガーで毎朝実行する想定。トリガーは手動で設定）
 */
function notifyUnchecked() {
  const list = apiUnchecked(CONFIG.ALERT_DAYS_AHEAD);
  if (!list.length) { console.log('未チェックなし'); return 0; }
  const to = CONFIG.ALERT_TO || Session.getEffectiveUser().getEmail();
  const lines = list.map(a => a.ship + '  ' + a.z + ' ' + a.g + '号機  ' + (a.to || '') + (a.status === '否' ? '  ★否（処置待ち）' : ''));
  const url = ScriptApp.getService().getUrl() || '';
  MailApp.sendEmail(to, '【アーム】出荷前タップ未チェック ' + list.length + '本',
    '今日〜' + CONFIG.ALERT_DAYS_AHEAD + '日後に出荷するアームで、ねじ穴タップ「良」の記録がないもの：\n\n' + lines.join('\n') + (url ? '\n\n記録はこちら: ' + url : ''));
  console.log('通知: %s 本 → %s', list.length, to);
  return list.length;
}
