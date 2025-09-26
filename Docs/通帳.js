// スクリプトプロパティに 'GEMINI_API_KEY' という名前でAPIキーを保存してください。
// プロジェクト番号　ｂ
// A present from ENJOINT group

/**
 * WebアプリとしてアクセスされたときにHTMLサービスを提供します。
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} HTML出力オブジェクト
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index') // HTMLファイル名を 'Index.html' と想定
      .setTitle('通帳データ化')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://drive.google.com/uc?id=1tK1IsXWIWvr-DzQ2mqWvvi0lxNewKg-5&.png');
}

/**
 * 抽出された取引データの残高検算を行います。
 * @param {Array<Object>} transactions - 抽出された取引データの配列。
 * @return {boolean} 検算結果が全て一致すればtrue、不一致があればfalse。
 */
function verifyBalances(transactions) {
  if (!transactions || transactions.length < 1) {
    return true; // データがない場合は検証不要
  }

  for (let i = 0; i < transactions.length; i++) {
    const currentTx = transactions[i];
    if (typeof currentTx.残高 !== 'number' || 
        (currentTx.出金額 !== null && typeof currentTx.出金額 !== 'number') ||
        (currentTx.入金額 !== null && typeof currentTx.入金額 !== 'number')) {
      console.warn("残高または金額が数値でないため、検算をスキップ:", currentTx);
      continue; 
    }
    
    if (i === 0) { 
        if (currentTx.取引内容 && currentTx.取引内容.includes("繰越")) {
            // 繰越行はそのまま信用
        }
        continue;
    }

    const previousTx = transactions[i-1];
    if (typeof previousTx.残高 !== 'number') {
        console.warn("前回の残高が数値でないため、検算をスキップ:", previousTx);
        continue; 
    }

    const expectedBalance = previousTx.残高 + (currentTx.入金額 || 0) - (currentTx.出金額 || 0);
    
    if (Math.round(currentTx.残高) !== Math.round(expectedBalance)) { 
      console.error(`残高不一致: 行 ${i+1} (${currentTx.取引日})`);
      console.error(`  前日残高: ${previousTx.残高}, 入金額: ${currentTx.入金額 || 0}, 出金額: ${currentTx.出金額 || 0}`);
      console.error(`  期待される残高: ${expectedBalance}, 実際の残高: ${currentTx.残高}`);
      return false; 
    }
  }
  return true; 
}


/**
 * クライアントサイドから呼び出され、画像データをGemini APIで処理します。
 * 不一致があった場合は1度だけ再分析を試みます。
 * @param {string} imageDataBase64 画像のBase64エンコード文字列
 * @param {string} mimeType 画像のMIMEタイプ
 * @param {boolean} includeHandwriting 手書き文字を含めるかどうかのフラグ
 * @return {Object} Gemini APIからのレスポンスオブジェクト、またはエラーオブジェクト
 */
function callGeminiApi(imageDataBase64, mimeType, includeHandwriting) { // 関数名を以前動作した名前に変更
  let analysisAttempt = 1;
  let lastError = null;

  while (analysisAttempt <= 2) { 
    try {
      const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!apiKey) {
        console.error("APIキーがスクリプトプロパティに設定されていません。'GEMINI_API_KEY'という名前で設定してください。");
        throw new Error("サーバーエラー: APIキーが設定されていません。");
      }

      const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

      let handwritingInstruction = includeHandwriting ? "手書きの文字や数字も認識に含めてください。" : "手書きと思われる文字や数字は無視し、印字された文字を中心に認識してください。";
      
      const currentGregorianYear = new Date().getFullYear(); 
      const reiwaStartYear = 2019; 
      const heiseiStartYear = 1989; 
      const showaStartYear = 1926;  
      const currentReiwaYear = currentGregorianYear - reiwaStartYear + 1; 

      let reAnalysisInstruction = "";
      if (analysisAttempt > 1) {
        reAnalysisInstruction = "前回、抽出された取引データの残高計算に不一致がありました。特に金額（出金額、入金額、残高）と日付の読み取り精度を向上させ、各行の計算が前行の残高と整合するように、より注意深く数値を読み取ってください。";
        console.log("再分析を試行します。追加指示:", reAnalysisInstruction);
      }
      
      // 濁点・半濁点に関する指示を強化
      const dakutenInstruction = `日本語の文字認識、特に濁点（゛）や半濁点（゜）の識別は非常に重要です。
例えば、「シ」と「ジ」、「ハ」と「バ」と「パ」、「カ」と「ガ」、「タ」と「ダ」などを正確に見分けてください。
具体例として、「ハカタゼイムショ」が「ハカタセイムショ」のように誤読されることがあります。このような誤りを防いでください。
特に固有名詞や金融機関名、摘要欄の詳細な記述において、これらの違いが重要です。細心の注意を払ってください。`;

      const prompt = `この通帳の画像から取引明細を抽出してください。画像の最下部まで、全ての行を注意深く読み取ってください。
${dakutenInstruction}
${reAnalysisInstruction}
以下のJSONスキーマに厳密に従って結果を返してください。
各取引について、取引日（yyyy-mm-dd形式、不明な場合はnull）、出金額（半角整数、該当なければ0）、入金額（半角整数、該当なければ0）、残高（半角整数、不明な場合はnull）、取引内容（文字列、摘要など、不明な場合は空文字）を抽出してください。取引内容が複数行にわたる場合は、各行をスペースで連結して1つの文字列としてください。

日付の年は西暦 (yyyy-mm-dd形式) でお願いします。
現在の西暦年は ${currentGregorianYear}年 (令和${currentReiwaYear}年) です。
通帳の年は和暦の元号 (例: R, H, S, 令, 平, 昭) に続く数字 (例: R6, H4, S60) や、元号が省略された数字のみ (例: 06, 30, 4) で記載されている場合があります。

1.  元号と共に年が記載されている場合 (例: R6年, H4年, 昭60年):
    それを正確に西暦に変換してください。
    (令和 (R, 令): ${reiwaStartYear}年が元年。例: R6は${reiwaStartYear + 6 - 1}年)
    (平成 (H, 平): ${heiseiStartYear}年が元年。例: H4は${heiseiStartYear + 4 - 1}年、H30は${heiseiStartYear + 30 - 1}年)
    (昭和 (S, 昭): ${showaStartYear}年が元年。例: S60は${showaStartYear + 60 - 1}年)

2.  数字のみで年が記載されている場合 (例: 06, 05, 04, 63, 2):
    - まず、その数字が現在の令和の年 (${currentReiwaYear}や${currentReiwaYear-1}など) に近いかどうかを確認します。
      例: 「${String(currentReiwaYear-1).padStart(2,'0')}」なら令和${currentReiwaYear-1}年 (${currentGregorianYear-1}年)、「${String(currentReiwaYear-2).padStart(2,'0')}」なら令和${currentReiwaYear-2}年 (${currentGregorianYear-2}年) のように、直近の令和の年として解釈することを優先します。
    - 次に、通帳の他の日付との連続性や文脈を考慮します。例えば、画像内に「03」「04」「05」といった連続した年の記載があり、それらが現在の令和の年より小さい場合、これらは過去の連続した年 (例: 令和3年, 令和4年, 令和5年、または平成3年, 平成4年, 平成5年など) を示している可能性が高いです。
    - ご提示の画像例のように「04.06.27」や「05.02.20」といった日付が連続している場合、「04」は${reiwaStartYear + 4 - 1}年 (令和4年) または ${heiseiStartYear + 4 - 1}年 (平成4年) など、文脈に合う年としてください。同様に「05」は ${reiwaStartYear + 5 - 1}年 (令和5年) または ${heiseiStartYear + 5 - 1}年 (平成5年) などと解釈してください。前後の日付との整合性を最も重視してください。
    - 非常に大きな数字 (例: 60, 63) の場合は、昭和の年である可能性が高いです (例: 60は昭和60年 ${showaStartYear + 60 - 1}年)。
    - どうしても判断が難しい場合は、令和の最も近い過去の年としてください。

3.  年が1桁の場合 (例: 4年、5年) は、元号が省略されている可能性が高いので、上記2のルールと文脈から判断してください。
4.  年が不明な場合はnullとしてください。
月日はゼロ埋めされた2桁で記載されていることが多いです (例: 06月07日)。

金額が「***」や「---」のようにマスクされている場合は0としてください。
繰り越し行など、出金額と入金額が両方とも0になるような実質的な取引ではない行は抽出対象外としてください。
${handwritingInstruction}
`;

      const payload = {
        contents: [{
          role: "user",
          parts: [
            { text: prompt },
            { inlineData: { mimeType: mimeType, data: imageDataBase64 } }
          ]
        }],
        generationConfig: {
          responseMimeType: "application/json",
          responseSchema: {
            type: "ARRAY",
            items: {
              type: "OBJECT",
              properties: {
                "取引日": { "type": "STRING", "description": "取引日 (yyyy-mm-dd形式、不明な場合はnull)" },
                "出金額": { "type": "NUMBER", "description": "出金額 (半角整数、該当なければ0)" },
                "入金額": { "type": "NUMBER", "description": "入金額 (半角整数、該当なければ0)" },
                "残高": { "type": "NUMBER", "description": "残高 (半角整数、不明な場合はnull)" },
                "取引内容": { "type": "STRING", "description": "取引内容 (摘要など、不明な場合は空文字、複数行の場合はスペースで連結)" }
              },
              required: ["取引日", "出金額", "入金額", "残高", "取引内容"]
            }
          }
        }
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      console.log(`Attempt ${analysisAttempt}: Gemini API Request Payload (first 700 chars of prompt):`, prompt.substring(0,700) + "..."); // ログ出力文字数を少し増やしました

      const response = UrlFetchApp.fetch(apiUrl, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();

      console.log(`Attempt ${analysisAttempt}: Gemini API Response Code:`, responseCode);

      if (responseCode === 200) {
        let parsedTransactions;
        try {
          const apiResult = JSON.parse(responseBody);
          if (apiResult.candidates && apiResult.candidates[0] && apiResult.candidates[0].content && apiResult.candidates[0].content.parts && apiResult.candidates[0].content.parts[0]) {
             const textResult = apiResult.candidates[0].content.parts[0].text;
             parsedTransactions = JSON.parse(textResult); 
             if (!Array.isArray(parsedTransactions)) {
                 throw new Error("APIからの応答が期待される配列形式ではありません。");
             }
          } else {
            throw new Error("API応答の構造が予期しないものです。");
          }

        } catch (e) {
          console.error(`Attempt ${analysisAttempt}: Gemini APIレスポンスのJSONパースに失敗:`, e.toString(), responseBody);
          throw new Error("APIレスポンスの解析に失敗しました。");
        }
        
        parsedTransactions = parsedTransactions.filter(item => !( (typeof item.出金額 === 'number' && item.出金額 === 0) && (typeof item.入金額 === 'number' && item.入金額 === 0) ));

        if (verifyBalances(parsedTransactions)) {
          console.log(`Attempt ${analysisAttempt}: 残高検算成功。`);
          return JSON.parse(responseBody); 
        } else {
          console.warn(`Attempt ${analysisAttempt}: 残高検算失敗。`);
          if (analysisAttempt < 2) {
            lastError = new Error("残高の整合性が取れませんでした。再分析を試みます。"); 
          } else {
            console.error("再分析後も残高検算失敗。最終結果として返します。");
            return JSON.parse(responseBody); 
          }
        }
      } else {
        console.error(`Attempt ${analysisAttempt}: Gemini API Error. Code: ${responseCode}. Body (first 500 chars): ${responseBody.substring(0,500)}`);
        let errorMessage = `Gemini APIとの通信に失敗しました (コード: ${responseCode})。`;
        try {
          const errorResponse = JSON.parse(responseBody);
          if (errorResponse && errorResponse.error && errorResponse.error.message) {
            errorMessage += ` 詳細: ${errorResponse.error.message}`;
          }
        } catch (e) {
          errorMessage += ` 応答 (一部): ${responseBody.substring(0, 200)}`;
        }
        throw new Error(errorMessage);
      }
      analysisAttempt++;
    } catch (error) {
      console.error(`callGeminiApi 関数内でエラー (Attempt ${analysisAttempt}):`, error.toString(), error.stack); // 関数名を修正
      lastError = error; 
      if (analysisAttempt < 2 && !(error.message.includes("APIキーが設定されていません"))) { 
         analysisAttempt++; 
      } else {
        return { 
          error: { 
            message: (lastError ? lastError.message : null) || error.message || "サーバー側で不明なエラーが発生しました。" 
          } 
        };
      }
    }
  }
   return { 
      error: { 
        message: (lastError ? lastError.message : null) || "不明な処理エラーが発生しました（ループ終了）。"
      } 
    };
}

