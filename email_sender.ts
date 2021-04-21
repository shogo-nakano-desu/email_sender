// MailApp.sendEmail(emailAddress, subject, main)
// spreadsheet ID を入力
const ssid:string  = "*****************",
        ss = SpreadsheetApp.openById(ssid),
        sheet = ss.getSheetByName("********"); //sheetname

function emailSender() {
    // 値が存在するデータ範囲を取得する
    const dataRange = sheet.getDataRange();

    // 列数を取得（列数固定だと、列数が変わったときに機能しなくなるため）
    let nameIndex:number = 0;
    let emailIndex:number = 0;
    let subjectIndex:number = 0;
    let bodyIndex:number = 0;
    //getNumColumns()→セル範囲にある列の列数を取得する
    for (let i:number = 1; i <= dataRange.getNumColumns(); i++) {
        switch (sheet.getRange(1, i).getValue()) {
        case 'email':
            emailIndex = i -1;// 二次元配列のループは0から始まるため、列数を -1 する（以下同様）
        case 'name':
            nameIndex = i -1;
        case 'subject':
            subjectIndex = i -1;
        default:
            ;
        }
    }

    // E-mail送信
    const data = dataRange.getValues();
    for (var i = 1; i < dataRange.getNumRows(); i++) {
        let name:string = data[i][nameIndex];  // 0行目はヘッダー情報のためスキップ（以下同様）
        let email:string = data[i][emailIndex];
        let subject:string = data[i][subjectIndex];
        // let body:string = data[i][bodyIndex];

        // メール本文(この形でシートの中に本文を入れる)
        let body = `
        ${name}様

        お世話になっております。

        この度、${name}様の率直なお気持ち、ご意見を頂戴したく、
        アンケートにご協力頂きたく存じます。

        何卒宜しくお願い申し上げます。
        `

        try {
        GmailApp.sendEmail(email, subject, body, {from: '*****@gmail.com', bcc: '*****@gmail.com'});
        console.log('送信OK: ${email} ${name} ${subject} ${body}');
        } catch (e) {
        console.log('送信NG: ${name} ${e.name} ${e.message}');
        }
    }
}
