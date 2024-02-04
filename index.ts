// 使うデータの型だけ定義
type YwtResult = {
  properties: {
    Writer: {
      people?: {
        id: string;
      }[];
    };
    いいね: {
      people?: {
        id: string;
      }[];
    };
    Name: {
      title?: {
        plain_text: string;
      }[];
    };
  };
  url: string;
};

type MemberResult = {
  properties: {
    アカウント: {
      people?: {
        id: string;
      }[];
    };
    Name: {
      title?: {
        plain_text: string;
      }[];
    };
    社員番号: {
      number: number;
    };
    旧メンバー: {
      checkbox: boolean;
    };
  };
  ywtTotalScore: number;
  ywtCount: number;
  url: string;
};

type YwtData = {
  results: YwtResult[];
  has_more: boolean;
  next_cursor: string | null;
};

type MemberData = {
  results: MemberResult[];
  has_more: boolean;
  next_cursor: string | null;
};

const START_ROW = 2;
const START_COLUMN = 4;

function getColName(num: number) {
  let sheet = SpreadsheetApp.getActiveSheet();
  return sheet.getRange(1, num).getA1Notation().replace(/\d/, "");
}

// YWTデータを取得
const fetchYwtData = (option?: { filter: object }) => {
  const token = PropertiesService.getScriptProperties().getProperty(
    "TOKEN_YWT_GRASS_MAP"
  );
  const databaseId =
    PropertiesService.getScriptProperties().getProperty("DATABASE_ID_YWT");
  const url = `https://api.notion.com/v1/databases/${databaseId}/query`;

  const options = {
    headers: {
      Authorization: `Bearer ${token}`,
      "Notion-Version": "2022-06-28",
    },
    contentType: "application/json",
    method: "post" as const,
  };

  let results: YwtData["results"] = [];
  let nextCursor: YwtData["next_cursor"] = null;
  do {
    const data = JSON.parse(
      UrlFetchApp.fetch(url, {
        ...options,
        payload: JSON.stringify({
          filter: option?.filter,
          start_cursor: nextCursor ?? undefined,
        }),
      }).getContentText()
    ) as YwtData;
    results = [...results, ...data.results];
    nextCursor = data.next_cursor;
  } while (nextCursor);

  return results;
};

// メンバーデータを取得
const fetchMemberData = (option?: { filter: object }) => {
  const token = PropertiesService.getScriptProperties().getProperty(
    "TOKEN_YWT_GRASS_MAP"
  );
  const databaseId =
    PropertiesService.getScriptProperties().getProperty("DATABASE_ID_MEMBER");
  const url = `https://api.notion.com/v1/databases/${databaseId}/query`;

  const options = {
    headers: {
      Authorization: `Bearer ${token}`,
      "Notion-Version": "2022-06-28",
    },
    contentType: "application/json",
    method: "post" as const,
  };

  let results: MemberData["results"] = [];
  let nextCursor: MemberData["next_cursor"] = null;
  do {
    const data = JSON.parse(
      UrlFetchApp.fetch(url, {
        ...options,
        payload: JSON.stringify({
          filter: option?.filter,
          start_cursor: nextCursor ?? undefined,
        }),
      }).getContentText()
    ) as MemberData;
    results = [...results, ...data.results];
    nextCursor = data.next_cursor;
  } while (nextCursor);

  return results
    .filter((result) => {
      return (
        !result.properties.旧メンバー.checkbox &&
        result.properties.Name.title?.at(0)?.plain_text &&
        result.properties.アカウント.people?.at(0)?.id
      );
    })
    .sort((a, b) => {
      if (
        (a.properties.社員番号.number === null ||
          a.properties.社員番号.number === undefined) &&
        (b.properties.社員番号.number === null ||
          b.properties.社員番号.number === undefined)
      ) {
        return 0;
      }
      if (
        a.properties.社員番号.number === null ||
        a.properties.社員番号.number === undefined
      ) {
        return 1;
      }
      if (
        b.properties.社員番号.number === null ||
        b.properties.社員番号.number === undefined
      ) {
        return -1;
      }
      return a.properties.社員番号.number - b.properties.社員番号.number;
    });
};

// 日付データを生成
const createDateData = () => {
  const date = new Date();

  // 2023年1月1日からの今日までの日付の配列データを取得
  // フォーマットはyyyy-MM-dd
  const dateArray = [];

  let dayCount = 0;
  let targetDate = new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate() - dayCount
  );
  // 2023年1月1日より前なら終了
  while (targetDate.getTime() < new Date(2023, 0, 1).getTime()) {
    dateArray.push(
      Utilities.formatDate(targetDate, "Asia/Tokyo", "yyyy-MM-dd")
    );
    dayCount++;
    targetDate = new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate() - dayCount
    );
  }

  for (let i = 0; i < 365; i++) {
    const targetDate = new Date(
      date.getFullYear(),
      date.getMonth(),
      date.getDate() - i
    );

    // 2023年1月1日より前なら終了
    if (targetDate.getTime() < new Date(2023, 0, 1).getTime()) break;

    dateArray.push(
      Utilities.formatDate(targetDate, "Asia/Tokyo", "yyyy-MM-dd")
    );
  }

  return dateArray;
};

const init = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getName() !== "YWTマップ") {
    const ui = SpreadsheetApp.getUi();
    ui.alert("シート「YWTマップ」で実行してください");
    return;
  }

  const YwtData = fetchYwtData({
    filter: {
      timestamp: "created_time",
      created_time: {
        on_or_after: "2023-01-01",
      },
    },
  });
  const memberData = fetchMemberData();
  const dateData = createDateData();

  memberData.forEach((data) => {
    data.ywtTotalScore = 0;
    data.ywtCount = 0;
  });

  // シートをクリア
  sheet.clearContents();

  // YWTの配列を作成
  const memberLength = memberData.length;
  const dateLength = dateData.length;
  for (let row = 0; row < memberLength; row++) {
    const targetMember = memberData[row];
    for (let column = 0; column < dateLength; column++) {
      const targetDate = dateData[column];
      // 該当するYWTがある場合、カウント
      const targetYwt = YwtData.find(
        (Ywt) =>
          Ywt.properties.Name.title?.at(0)?.plain_text !== undefined &&
          (Ywt.properties.Name.title?.at(0)?.plain_text as string).startsWith(
            targetDate
          ) &&
          Ywt.properties.Writer.people?.at(0)?.id ===
            targetMember.properties.アカウント.people?.at(0)?.id
      );
      if (targetYwt) {
        const iineScore = targetYwt.properties.いいね.people?.length;
        const ywtScore = iineScore ? iineScore + 1 : 1;
        const ywtLink = targetYwt.url;
        targetMember.ywtTotalScore += ywtScore;
        targetMember.ywtCount += 1;
        const link = `=HYPERLINK("${ywtLink}", ${ywtScore})`;
        sheet
          .getRange(
            `${getColName(column + START_COLUMN + 1)}${row + START_ROW + 2}`
          )
          .setFormula(link);
      }
    }
  }

  // 行のメンバー一覧を描画
  sheet.getRange(`A2`).setValue("社員番号");
  sheet.getRange(`B2`).setValue("member");
  sheet.getRange(`C2`).setValue("総スコア");
  sheet.getRange(`D2`).setValue("提出数");
  for (let row = 0; row < memberLength; row++) {
    const rowMember = memberData[row];
    const memberName = rowMember.properties.Name.title?.at(0)?.plain_text;
    const memberLink = rowMember.url;
    const link = `=HYPERLINK("${memberLink}", "${memberName}")`;
    sheet
      .getRange(`${getColName(START_COLUMN - 3)}${row + START_ROW + 2}`)
      .setValue(rowMember.properties.社員番号.number);
    sheet
      .getRange(`${getColName(START_COLUMN - 2)}${row + START_ROW + 2}`)
      .setFormula(link);
    sheet
      .getRange(`${getColName(START_COLUMN - 1)}${row + START_ROW + 2}`)
      .setValue(rowMember.ywtTotalScore);
    sheet
      .getRange(`${getColName(START_COLUMN)}${row + START_ROW + 2}`)
      .setValue(rowMember.ywtCount);
  }
  // 列の日付一覧を描画
  for (let column = 0; column < dateLength; column++) {
    const targetDate = dateData[column];
    sheet
      .getRange(`${getColName(column + START_COLUMN + 1)}${START_ROW}`)
      .setValue(targetDate);
  }
  // 古いfilterを削除
  const oldFilter = sheet.getFilter();
  if (oldFilter) oldFilter.remove();
  // filterを新規設定
  const filter = sheet
    .getRange(
      START_ROW + 1,
      1,
      memberLength + START_ROW,
      dateLength + START_COLUMN
    )
    .createFilter();
  filter.sort(START_COLUMN - 1, false);

  const date = new Date();
  sheet
    .getRange(`A1`)
    .setValue(
      "更新日時：" +
        Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")
    );
};

const onOpen = () => {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("YWTマップ設定", [
    { name: "表を更新", functionName: "init" },
  ]);
};
