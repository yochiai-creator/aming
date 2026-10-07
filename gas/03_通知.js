/**
 * 出荷前の未チェック通知（時間主導トリガーで毎朝実行する想定。トリガーは手動で設定）
 * 次の出荷日（翌営業日）に出荷するアームのうち、ねじ穴タップ「良」の記録がないものだけを送る
 */
function notifyUnchecked() {
  const day = nextShipDate_();
  if (!day) { console.log('次の出荷日なし'); return 0; }
  const list = uncheckedBetween_(day, day);
  if (!list.length) { console.log('%s 出荷分は全部チェック済み', day); return 0; }
  const to = cfg_('ALERT_TO') || Session.getEffectiveUser().getEmail();
  const lines = list.map(a => a.z + ' ' + a.g + '号機  ' + (a.to || '') + (a.status === '否' ? '  ★否（処置待ち）' : ''));
  const url = ScriptApp.getService().getUrl() || '';
  MailApp.sendEmail(to, '【aming】' + day.slice(5) + ' 出荷分 タップ未チェック ' + list.length + '本',
    day + ' に出荷するアームで、ねじ穴タップ「良」の記録がないもの：\n\n' + lines.join('\n') + (url ? '\n\n記録はこちら: ' + url : ''));
  console.log('通知: %s 出荷分 %s 本 → %s', day, list.length, to);
  return list.length;
}
