/**
 * 職員室テレビ表示案（研究部会 提案資料）を、学校のGoogleアカウントで
 * 見られるWebページとして配信するだけの、単独のスクリプトです。
 * 西原中ダッシュボード本体（別プロジェクト）には一切手を加えません。
 *
 * 使い方は gas/tv-proposal-standalone/README.md を参照してください。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Proposal')
    .setTitle('職員室テレビ表示案')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
