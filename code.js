// =====================================
// 設定値
// =====================================
// ★★★ 転記先の新しいスプレッドシートのIDをここに入力してください！ ★★★
// 例: https://docs.google.com/spreadsheets/d/1234567890abcdefghijklmnopqrstuvwxyz/edit#gid=0 の斜線部分
const TARGET_SPREADSHEET_ID = "ここに転記先のスプレッドシートIDを記述"; 

const SHEET_NAME = "経費承認表"; // 転記先のシート名
const ADMIN_EMAIL = "admin@example.com"; // 管理者（CC）のメールアドレス
// 部署ごとの承認者メールアドレス
const approverSettings = {
  "営業部": "approver@example.com",
  "経理部": "acc_approver@example.com",
  "総務部": "gen_approver@example.com",
  "人事部": "hr_approver@example.com",
  "法務部": "legal_approver@example.com",
  "情報システム部": "it_approver@example.com",
  "その他": "other_approver@example.com",
};

// =====================================
// メイン処理
// =====================================
/**
 * フォーム送信時にトリガーされる関数
 * @param {Object} e フォーム送信イベントオブジェクト
 */
function onFormSubmit(e) {
  
  // 1. 転記先のスプレッドシートを開く
  let ss;
  try {
    ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  } catch (err) {
    Logger.log(`エラー: スプレッドシートID "${TARGET_SPREADSHEET_ID}" のファイルを開けませんでした。IDが正しいか、スクリプト実行者にアクセス権があるか確認してください。`);
    return;
  }
  
  const approvalSheet = ss.getSheetByName(SHEET_NAME);

  if (!approvalSheet) {
    Logger.log(`致命的なエラー: 転記先ファイル内にシート名 "${SHEET_NAME}" が見つかりません。`);
    return;
  }

  // フォームから送信された最新行のデータを取得 (フォーム連携シートからではないため e.values をそのまま使用)
  const data = e.values;

  // フォームの項目とe.valuesのインデックス（変更なし）
  const rawTimestamp = data[0];
  const timestamp = Utilities.formatDate(rawTimestamp, Session.getScriptTimeZone(), 'yyyy年MM月dd日');
  const applicantEmail = data[1];
  const applicantName = data[2];
  const department = data[3];
  const expenseType = data[4];
  const amount = data[5]; 
  const usageDate = data[6];
  const reason = data[7];
  const receiptUrl = data[8] ? data[8].split(',')[0] : 'なし'; 

  // 1. 承認用スプレッドシートにデータを転記
  const status = "未承認";
  
  // 承認用シートの列順にデータを再構成
  // 申請日(A) | 申請者名(B) | メールアドレス(C) | 所属部署(D) | 経費の種類(E) | 金額(F) | 利用日(G) | 内容・申請理由(H) | 
  // 領収書URL(I) | 承認ステータス(J) | 承認者コメント(K) | 承認者(L)
  const newData = [
    timestamp,
    applicantName,
    applicantEmail,
    department,
    expenseType,
    amount,    
    usageDate, 
    reason,
    receiptUrl,
    status,     
    "",         
    ""          
  ];
  
  approvalSheet.appendRow(newData);

  // 2. 承認者と管理者にメールを送信
  const approverEmail = approverSettings[department];

  if (!approverEmail) {
    Logger.log(`エラー: ${department} の承認者が見つかりません。 approverSettingsに設定を追加してください。`);
    return;
  }

  // 承認用シートの最終行URLを取得（転記された行）
  const lastRow = approvalSheet.getLastRow();
  const ssUrl = ss.getUrl();
  // 転記先ファイルへのリンクを作成
  const approvalLink = `${ssUrl}#gid=${approvalSheet.getSheetId()}&range=A${lastRow}`;

  // メール件名と本文
  const subject = `【経費申請】${applicantName}：${expenseType} (${department})`;
  const body = `
  以下の内容で経費申請がありました。承認をお願いいたします。

  ■ 申請概要
  ------------------------------------
  申請者: ${applicantName}
  所属部署: ${department}
  利用日: ${usageDate}  
  経費の種類: ${expenseType}
  金額: ${amount} 円
  内容・申請理由: ${reason}
  領収書: ${receiptUrl}
  ------------------------------------

  ■ 承認はこちらから
  ${approvalLink}

  ※ このメールはシステムからの自動送信です。
  `;

  // メール送信
  MailApp.sendEmail({
    to: approverEmail, // 承認者 (TO)
    cc: ADMIN_EMAIL, // 管理者のみにCC
    subject: subject,
    body: body.trim() 
  });

  Logger.log(`経費申請を処理しました。承認者: ${approverEmail}`);
}
