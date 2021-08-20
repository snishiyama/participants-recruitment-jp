// Copyright (c) 2017-2019, Sho Ishiguro (c) 2019-2021, Satoru Nishiyama and Sho Ishiguro
// Use of this source code is governed by the BSD 2-Clause License

const TYPE = 3; // 1: 自由回答, 2 or 3: 選択式 どちらかの半角数字を入れてください。

function init() {
  settings.init();
}

///////////////////////////////////////////////////////////////////////////////
// トリガー用の関数
///////////////////////////////////////////////////////////////////////////////

function onOpening() {
  SpreadsheetApp.getUi().createMenu('カレンダー').addItem('カレンダーをシートに反映', 'onCalendarUpdated').addToUi();
  if (sheets.length > 1) {
    mail.alertFewMails();
  }
}

function onFormSubmission(e) {
  try {
    // systemを利用しないなら以降の処理を行わない
    if (settings.config.useFormSystem != 1) {
      return;
    }
    // 実際の回答に続けて値のない回答が送られることがあるので以下のif文で回避
    if (e.values[settings.config.colAddress].length > 0 && !settings.isDefault()) {
      booking.values = e.values;
      booking.setEventType(ScriptApp.EventType.ON_FORM_SUBMIT).validate().allocate(e.range.getRow());
      const { name, address, from: fromWhen, to: toWhen, trigger } = booking;
      // mail
      mail.create(name, trigger, fromWhen, toWhen).setBcc('', settings.config.selfBccTentative).send(address).alertFewMails();
      onCalendarUpdated();
      console.log('SUCCESS!');
    } else {
      console.log(e.values);
    }
  } catch (err) {
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    MailApp.sendEmail(settings.config.experimenterMailAddress, 'エラーが発生しました', msg);
  }
}

function onSheetEdit(e) {
  try {
    const sh = e.range.getSheet();
    const sheetName = sh.getSheetName();
    // 「フォームの回答」シートが編集された場合
    if (sheetName === sheets.sheets[0].getSheetName()) {
      const srow = e.range.getRow();
      const erow = e.range.getLastRow();
      const scol = e.range.getColumn();
      // 予約ステータスの列を編集した場合
      if (scol === settings.config.colStatus + 1) {
        const answers = sh.getDataRange().getValues();
        for (let row = srow; row <= erow; row++) {
          // トリガーが削除された場合以降の処理をしない
          if (answers[row - 1][settings.config.colStatus] == '') {
            continue;
          }
          booking.values = answers[row - 1];
          // まだ予約に関するメールが送信されていない場合
          if (booking.values[settings.config.colMailed] !== 1) {
            booking.setEventType(ScriptApp.EventType.ON_EDIT).validate().allocate(row);
            const { name, address, from: fromWhen, to: toWhen, trigger, assistant } = booking;
            mail.create(name, trigger, fromWhen, toWhen).setBcc(assistant, settings.config.selfBccTentative).send(address).alertFewMails();
          }
        }
      }
    } else {
      // 「フォームの回答」以外のシートが編集された場合
      if (sheetName == '設定') {
        oldConfig = copy(settings.config); // すぐあとの比較のために古い設定をコピー
        settings.collect(sheetName, true).save(); // 更新・保存
        if (settings.config.remindHour != oldConfig.remindHour) {
          scriptTriggers.updateClockTrigger(settings.config.remindHour, settings.config.expTimeZone);
        } else if (settings.config.expTimeZone != oldConfig.expTimeZone) {
          sheets.ss.setSpreadsheetTimeZone(settings.config.expTimeZone);
          scriptTriggers.updateClockTrigger(settings.config.remindHour, settings.config.expTimeZone);
        } else if (settings.config.workingCalendar != oldConfig.workingCalendar) {
          schedule.calendar = CalendarApp.getCalendarById(settings.config.workingCalendar);
          // scriptTriggers.updateCalendarTrigger(settings.config.workingCalendar);
          alertInitWithChangeOf('参照するカレンダー');
        } else if (settings.config.experimentLength != oldConfig.experimentLength) {
          alertInitWithChangeOf('実験の所要時間');
        } else if (fmtDate(settings.config.openTime, 'HH:mm') != fmtDate(new Date(oldConfig.openTime), 'HH:mm')) {
          alertInitWithChangeOf('実験の開始時刻');
        } else if (fmtDate(settings.config.closeTime, 'HH:mm') != fmtDate(new Date(oldConfig.closeTime), 'HH:mm')) {
          alertInitWithChangeOf('実験の終了時刻');
        }
      } else if (sheetName == 'メンバー' || sheetName == 'テンプレート') {
        settings.collect(sheetName, true).save(); // 新しいキャッシュを作成
      } else if (sheetName == '空き予定') {
        onCalendarUpdated();
      }
    }
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
    // Browser.msgBox('エラーが発生しました', msg, Browser.Buttons.OK);
  }
}

function onClock() {
  try {
    // リマインダーの送信
    const answers = sheets.sheets[0].getDataRange().getValues();
    const timeNow = new Date();
    const tomorrowExps = [];
    for (let row = 0; row < answers.length; row++) {
      let ans = answers[row];
      if (ans[settings.config.colReminded] == '送信準備') {
        const remindDatetime = ans[settings.config.colRemindDate];
        if (is(remindDatetime, 'Date') && remindDatetime <= timeNow) {
          booking.values = ans;
          booking.setEventType(ScriptApp.EventType.CLOCK).allocate(row + 1);
          const { name, address, from: fromWhen, to: toWhen, assistant } = booking;
          mail.create(name, 'リマインダー', fromWhen, toWhen).setBcc(assistant, settings.config.selfBccReminder).send(address);
          tomorrowExps.push({ name: name, from: fromWhen, to: toWhen });
        }
      }
    }
    // 自分にもリマインダーを送る場合
    if (tomorrowExps.length > 0 && settings.config.sendTmrwExps > 0) {
      const tomorrow = new Date();
      tomorrow.setDate(new Date().getDate() + 1);
      tomorrowExps.sort((a, b) => {
        return a.from < b.from ? -1 : 1;
      });
      const time_table = tomorrowExps.map((exp) => {
        return `${fmtDate(exp.from, 'HH:mm')} - ${fmtDate(exp.to, 'HH:mm')} ${exp.name}`;
      });
      const body = time_table.join('\n');
      const title = `明日（${fmtDate(tomorrow, 'MM/dd')}）の実験予定`;
      MailApp.sendEmail(settings.config.experimenterMailAddress, title, body);
    }

    // フォームの修正
    if (settings.config.useFormSystem == 1) {
      form.modify();
    }
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    MailApp.sendEmail(settings.config.experimenterMailAddress, 'エラーが発生しました', msg);
  }
}

function onCalendarUpdated() {
  try {
    // type 3の時だけ動作させる
    if (TYPE != 3) {
      return;
    }
    schedule.update().allocate(); // スケジュールを更新してシートに反映する
    form.modify();
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
  }
}

///////////////////////////////////////////////////////////////////////////////
// Utility functions
///////////////////////////////////////////////////////////////////////////////

// 型判定のための関数https://qiita.com/Layzie/items/465e715dae14e2f601de より
function is(obj, type) {
  const clas = Object.prototype.toString.call(obj).slice(8, -1);
  return obj !== undefined && obj !== null && clas === type;
}

// 全角を半角に変換する関数
function zenToHan(str) {
  if (is(str, 'String')) {
    return str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
      // 全角を半角に変換
      return String.fromCharCode(s.charCodeAt(0) - 65248); // 10進数の場合
    });
  } else {
    return str;
  }
}

function numToColumnNotation(num) {
  const alphabet_upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let dgt = Math.floor(num / alphabet_upper.length);
  let remain = num % alphabet_upper.length;
  if (dgt < 1) {
    return alphabet_upper[remain];
  }
  return numToColumnNotation(dgt - 1) + alphabet_upper[remain];
}

function columnNotationToNum(notation) {
  const alphabet_upper = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  if (notation.length < 1) {
    return -1;
  } else if (notation.length == 1) {
    return alphabet_upper.indexOf(notation);
  }
  return (columnNotationToNum(notation.slice(0, -1)) + 1) * alphabet_upper.length + alphabet_upper.indexOf(notation.slice(-1));
}

function fmtDate(datetime, pattern) {
  if (is(datetime, 'Date')) {
    if (/yobi/.test(pattern)) {
      var yobi = new Array('日', '月', '火', '水', '木', '金', '土')[datetime.getDay()];
      pattern = pattern.replace(/yobi/, yobi);
    }
    return Utilities.formatDate(datetime, settings.config.expTimeZone, pattern);
  }
  return datetime;
}

function copy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

// https://qiita.com/jz4o/items/d4e978f9085129155ca6 を改変
function isHoliday(time) {
  //土日か判定
  let weekInt = time.getDay();
  if (weekInt <= 0 || 6 <= weekInt) {
    return true;
  }

  //祝日か判定
  const calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  if (calendar == null) {
    let msg = '祝日のカレンダーがご自身のgoogleカレンダーに登録されていません。実験日の休日判定を行う場合は祝日カレンダーを登録してください。';
    msg += 'もし休日判定を行わない場合は「テンプレート」シートのB列の数字をすべて0に変更してください。';
    msg += '\n\nなおこのエラーが発生したため参加者にはメールは送られていません。休日判定のための設定したのち';
    msg += '予約ステータスを含む右4列の内容を削除し再度予約ステータスにトリガーを入力してください。';
    throw new Error(msg);
  }
  const todayEvents = calendar.getEventsForDay(time);

  return todayEvents.length > 0;
}

function alertInitWithChangeOf(changed) {
  if (TYPE != 3) {
    return;
  }
  const choice = dlg.alert(
    `${changed}が変更されました`,
    '空き予定を初期化しますか？\n\n処理に時間がかかります。10〜20秒ほどお待ち下さい。 \n処理が完了するとその旨のダイアログボックスが表示されます。',
    dlg.ui.ButtonSet.OK_CANCEL
  );
  if (choice == dlg.ui.Button.OK) {
    schedule.init();
    form.modify();
    dlg.alert('空き予定の初期化', '空き予定の初期化が終了しました。適宜情報を変更してください。', dlg.ui.ButtonSet.OK);
  }
}

///////////////////////////////////////////////////////////////////////////////
// Objects
///////////////////////////////////////////////////////////////////////////////

// シートをいい感じに扱いやすくしてくれるはずのオブジェクト
const sheets = (function () {
  const __ss = SpreadsheetApp.getActiveSpreadsheet(); // spreadsheet
  let __sheets = __ss.getSheets();
  let __name_idx = new Map();
  __sheets.forEach((sh, idx) => __name_idx.set(sh.getName(), idx));

  let __values = { 設定: undefined, テンプレート: undefined, メンバー: undefined, 空き予定: undefined, Cached: undefined };

  return {
    get length() {
      return __sheets.length;
    },
    get sheets() {
      return __sheets;
    },
    get ss() {
      return __ss;
    },

    update: function () {
      __sheets = __ss.getSheets();
      __name_idx = new Map();
      __sheets.forEach((sh, idx) => __name_idx.set(sh.getName(), idx));
      return this;
    },

    getSheetByName: function (name) {
      if (__name_idx.get(name) === undefined) {
        new Error(`「${name}」シートがありません。`);
      }
      return __sheets[__name_idx.get(name)];
    },

    getValuesOf: function (name, update = false) {
      if (__values[name] === undefined || update) {
        const sh = this.getSheetByName(name);
        __values[name] = sh.getDataRange().getValues();
      }
      return __values[name];
    },

    getValueAt: function (sheetName, row, col) {
      const values = this.getValuesOf(sheetName);
      return values[row][col];
    },

    getTargetRowID(sheetName, col, target) {
      let sheetValues = this.getValuesOf(sheetName);
      for (let row = 0; row < sheetValues.length; row++) {
        var rowValues = sheetValues[row];
        if (rowValues[col] == target) {
          return row + 1; // getRangeで使うことを想定しているので，+1する
        }
      }
      return undefined;
    },
  };
})();

// 設定をいい感じに扱いやすくしてくれるはずのオブジェクト
const settings = (function () {
  let __settings = {};
  const __name_to_key = new Map([
    ['設定', 'config'],
    ['テンプレート', 'templates'],
    ['メンバー', 'members'],
  ]);

  function __getConfig(update) {
    const table = sheets.getValuesOf('設定', update);
    __settings.config = {};
    for (let row = 1; row < table.length; row++) {
      let key = table[row][1];
      let val = zenToHan(table[row][2]); // 念の為
      if (key.indexOf('col') == 0) {
        const is_alphabet = new RegExp(/^[a-zA-Z]*$/);
        if (!is_alphabet.test(val)) {
          throw new Error(
            `${key}に英字以外が入力されています。列に関する設定には英字を入力してください。一見，英字を入力しているにもかかわらずこのエラーが表示される場合は，入力内容にスペースが含まれているかもしれません。`
          );
        }
        val = columnNotationToNum(val.toUpperCase()); // 列番号に関する設定は，Numberに変更しておく
      }
      __settings.config[key] = val;
    }
    __arrangeExpPeriod();
  }

  function __getMailTemplates(update) {
    const table = sheets.getValuesOf('テンプレート', update);
    __settings.templates = {};
    for (let row = 1; row < table.length; row++) {
      let key = table[row][0];
      let property = {};
      property.changeByDay = table[row][1];
      property.title = table[row][2];
      property.bodywd = table[row][3];
      property.bodywe = table[row][4];
      __settings.templates[key] = property;
    }
  }

  function __getMembers(update) {
    const table = sheets.getValuesOf('メンバー', update);
    __settings.members = {};
    for (let row = 1; row < table.length; row++) {
      let key = zenToHan(table[row][0]);
      let address = zenToHan(table[row][2]);
      __settings.members[key] = address;
    }
  }

  function __collect(sheetName, update) {
    switch (sheetName) {
      case '設定':
        __getConfig(update);
        break;
      case 'テンプレート':
        __getMailTemplates(update);
        break;
      case 'メンバー':
        __getMembers(update);
        break;
      default:
        throw new Error(`${sheetName} は「設定」「テンプレート」「メンバー」のいずれにも一致しません`);
    }
  }

  function __retrieve() {
    const sh = sheets.getSheetByName('Cached');
    const cache_json = sh.getRange(1, 1).getValue();
    if (cache_json.length < 10) {
      // cache用JSONが存在しない（適切ではない）場合
      for (let sheetName of __name_to_key.keys()) {
        __collect(sheetName);
      }
    } else {
      __settings = JSON.parse(cache_json);
      __arrangeExpPeriod();
    }
  }

  function __arrangeExpPeriod() {
    // parseしたままだと以下の2つがstringのままで機能しない。
    // シートから直接値を取得した場合は問題ないが，それほど処理速度に影響が出るとも思えないので，処理を分けない
    __settings.config.openDate = new Date(__settings.config.openDate);
    __settings.config.closeDate = new Date(__settings.config.closeDate);

    // openTime, closeTime
    let temp_date = new Date();
    if (is(__settings.config.openTime, 'String')) {
      __settings.config.openTime = new Date(__settings.config.openTime);
    } else if (is(__settings.config.openTime, 'Number')) {
      temp_date.setHours(__settings.config.openTime, 0, 0, 0);
      __settings.config.openTime = new Date(temp_date);
    }
    if (is(__settings.config.closeTime, 'String')) {
      __settings.config.closeTime = new Date(__settings.config.closeTime);
    } else if (is(__settings.config.closeTime, 'Number')) {
      temp_date.setHours(__settings.config.closeTime, 0, 0, 0);
      __settings.config.closeTime = new Date(temp_date);
    }
    if (__settings.config.openTime.toString() == 'Invalid Date') {
      throw new Error('開始時刻の設定が適切ではありません。"10:00" のような時間表記 あるいは "10" のような"時"だけを示す数値を入力してください。');
    } else if (__settings.config.closeTime.toString() == 'Invalid Date') {
      throw new Error('終了時刻の設定が適切ではありません。"10:00" のような時間表記 あるいは "10" のような"時"だけを示す数値を入力してください。');
    }

    // 実験開始日・終了日の調整
    __settings.config.outOfDate = false;
    const now = new Date();
    if (__settings.config.openDate < now) {
      __settings.config.openDate = now;
    }
    if (__settings.config.closeDate < now) {
      __settings.config.outOfDate = true;
    }
    // 実験開始日・終了日の日時の設定
    const openHour = __settings.config.openTime.getHours();
    const openMin = __settings.config.openTime.getMinutes();
    const closeHour = __settings.config.closeTime.getHours();
    const closeMin = __settings.config.closeTime.getMinutes();
    __settings.config.openDate.setHours(openHour, openMin, 0, 0);
    __settings.config.closeDate.setHours(closeHour, closeMin, 0, 0);
  }

  if (sheets.length > 1) {
    __retrieve();
  }

  return {
    get config() {
      if (__settings.config === undefined) {
        this.collect('設定');
      }
      return __settings.config;
    },
    get templates() {
      if (__settings.templates === undefined) {
        this.collect('テンプレート');
      }
      return __settings.templates;
    },
    get members() {
      if (__settings.members === undefined) {
        this.collect('メンバー');
      }
      return __settings.members;
    },

    collect: function (sheetName, update = false) {
      __collect(sheetName, update);
      return this;
    },

    retrieve: function () {
      __retrieve();
      return this;
    },

    save: function () {
      const sh = sheets.getSheetByName('Cached');
      sh.getRange(1, 1).setValue(JSON.stringify(__settings));
    },

    isDefault: function () {
      const is_default = {
        実験者名: this.config.experimenterName == '実験太郎',
        電話番号: this.config.experimenterPhone == 'xxx-xxx-xxx',
        実施場所: this.config.experimentRoom == '実施場所',
      };

      const title = '設定がデフォルトのままです';
      let msg = '以下の重要な設定がデフォルトのままだったので，参加希望者への予約確認メールの送信を中止しました。\n\n';
      let is_any_default = false;
      for (const key in is_default) {
        if (is_default[key]) {
          msg += `${key}\n`;
          is_any_default = true;
        }
      }

      if (is_any_default) {
        msg += '\n変更後，再度参加者応募のテストをして，予約確認のメールが送信されるかどうか，およびその本文が適切かどうかを確認してください。';
        MailApp.sendEmail(this.config.experimenterMailAddress, title, msg);
      }

      return is_any_default;
    },
  };
})();

// メールの内容やらを作成するためのオブジェクト
const mail = (function () {
  let __title;
  let __body;
  let __bcc;
  let __remaining;

  return {
    get remaining() {
      if (__remaining === undefined) {
        __remaining = MailApp.getRemainingDailyQuota();
      }
      const sh = sheets.getSheetByName('設定');
      const row_id = sheets.getTargetRowID('設定', 1, 'remainingMails');

      sh.getRange(row_id, 3).setValue(__remaining);
      return __remaining;
    },

    create: function (name, trigger, from, to) {
      const template = settings.templates[trigger];

      // タイトル
      __title = template.title;

      // 本文
      __body = template.bodywd;
      if (template.changeByDay == 1 && isHoliday(from)) {
        __body = template.bodywe;
      }
      settings.config.participantName = name;
      settings.config.expDate = fmtDate(from, 'MM/dd（yobi）');
      settings.config.fromWhen = fmtDate(from, 'HH:mm');
      settings.config.toWhen = fmtDate(to, 'HH:mm');
      settings.config.openDate = fmtDate(settings.config.openDate, 'yyyy/MM/dd');
      settings.config.closeDate = fmtDate(settings.config.closeDate, 'yyyy/MM/dd');
      // メールの本文の変数を置換する
      for (const key in settings.config) {
        let regex = new RegExp(key, 'g');
        __body = __body.replace(regex, settings.config[key]);
      }

      return this;
    },

    setBcc: function (assistants, selfBcc) {
      const bcc_array = [];
      if (selfBcc > 0) {
        bcc_array.push(settings.config.experimenterMailAddress);
      }
      assistants = String(zenToHan(assistants));
      // 担当が空欄でなければ
      if (assistants.length > 0) {
        const assistantIDs = assistants.match(/\d+/g);
        assistantIDs.forEach((ast_id) => bcc_array.push(settings.members[ast_id]));
      }
      __bcc = bcc_array.join(','); // 配列が空なら''が返される

      return this;
    },

    send: function (address) {
      if (__bcc.length > 5) {
        MailApp.sendEmail(address, __title, __body, { bcc: __bcc });
      } else {
        MailApp.sendEmail(address, __title, __body);
      }

      return this;
    },

    alertFewMails() {
      const thresholds = [5, 10, 20];
      if (thresholds.includes(this.remaining)) {
        const title = '自動送信メールの残数が' + String(this.remaining) + 'です。';
        const message =
          title +
          'この24時間以内に送信されるかもしれない予約の確認やリマインダーのメール数を考慮して予約を完了させてください。' +
          '自分や分担者にもメールが送信されるようにしている場合は1通あたりに減る数が 2, 3... 大きくなります。';
        dlg.alert(title, message, dlg.ui.ButtonSet.OK);
        // Browser.msgBox(title, message, Browser.Buttons.OK);
      }
    },
  };
})();

// フォームを扱うオブジェクト
const form = (function () {
  let __form;

  function __modifyType2() {
    if (__form === undefined) {
      __form = FormApp.openByUrl(sheets.ss.getFormUrl());
    }
    const items = __form.getItems();
    const itemForDate = items[settings.config.colExpDate - 1]; // -1 なのは，シートで フォームの送信時間が増えているから
    let item;
    if (itemForDate.getType() == 'LIST') {
      item = itemForDate.asListItem();
    } else if (itemForDate.getType() == 'MULTIPLE_CHOICE') {
      item = itemForDate.asMultipleChoiceItem();
    } else {
      return;
    }

    let firstDateOfChoices = new Date(settings.config.openDate);
    firstDateOfChoices.setHours(0, 0, 0, 0);
    settings.config.closeDate.setHours(0, 0, 0, 0);
    // 設定された実験の開始日が関数の動作日時よりも前の場合
    if (firstDateOfChoices < new Date()) {
      firstDateOfChoices.setDate(new Date().getDate() + 1);
    }
    const choices = [];
    for (const choiceDate = firstDateOfChoices; choiceDate <= settings.config.closeDate; choiceDate.setDate(choiceDate.getDate() + 1)) {
      const newChoice = item.createChoice(fmtDate(choiceDate, 'yyyy/MM/dd'));
      choices.push(newChoice);
    }
    item.setChoices(choices);
  }

  function __modifyType3() {
    if (__form === undefined) {
      __form = FormApp.openByUrl(sheets.ss.getFormUrl());
    }
    const items = __form.getItems();
    const itemForDate = items[settings.config.colExpDate - 1]; // -1 なのは，シートで フォームの送信時間が増えているから
    let item;
    if (itemForDate.getType() == 'LIST') {
      item = itemForDate.asListItem();
    } else if (itemForDate.getType() == 'MULTIPLE_CHOICE') {
      item = itemForDate.asMultipleChoiceItem();
    } else {
      return;
    }

    const choices = [];
    for (const exp_dates of schedule.available.values()) {
      for (let idx = 0; idx < exp_dates.length; idx++) {
        const exp_date = exp_dates[idx];
        if (is(exp_date, 'Date') && settings.config.openDate <= exp_date && exp_date < settings.config.closeDate) {
          const stime = fmtDate(new Date(exp_date), 'yyyy/MM/dd HH:mm');
          let etime = new Date(stime);
          etime.setMinutes(etime.getMinutes() + settings.config.experimentLength);
          etime = fmtDate(etime, 'HH:mm');
          choices.push(item.createChoice(`${stime}-${etime}`));
        }
      }
    }
    item.setChoices(choices);
  }

  return {
    modify: function () {
      // 実験実施期間を過ぎていたらフォームを閉じる
      if (settings.config.outOfDate) {
        if (__form === undefined) {
          __form = FormApp.openByUrl(sheets.ss.getFormUrl());
        }
        __form.setAcceptingResponses(false);
        return;
      }
      switch (TYPE) {
        case 2:
          return __modifyType2();
        case 3:
          return __modifyType3();
        default:
          return;
      }
    },
  };
})();

// 予約情報を扱うオブジェクト
const booking = (function () {
  let __values;
  let __name;
  let __address;
  let __from;
  let __to;
  let __valid = false;
  let __trigger;
  let __status;
  let __event_type;
  let __calendar;
  let __finalizeTriggers;
  if (sheets.length > 1) {
    __calendar = CalendarApp.getCalendarById(settings.config.workingCalendar);
    __finalizeTriggers = String(settings.config.finalizeTrigger).match(/\d+/g);
  }

  function __isValidDatetime() {
    if (__from === undefined || __to === undefined) {
      return false;
    }
    settings.config.openTime.setFullYear(__from.getFullYear(), __from.getMonth(), __from.getDate());
    settings.config.closeTime.setFullYear(__from.getFullYear(), __from.getMonth(), __from.getDate());
    const isValidTime = settings.config.openTime <= __from && __to <= settings.config.closeTime;
    const isValidDate = settings.config.openDate <= __from && __from <= settings.config.closeDate;
    return isValidTime && isValidDate;
  }

  function __validateSubmission() {
    if (__from === undefined || __to === undefined) {
      return undefined;
    }
    const events = __calendar.getEvents(__from, __to);
    __trigger = '仮予約';
    __status = ['', '', '', ''];
    __valid = true;
    if (events.length > 0) {
      __trigger = '重複';
      __status = [__trigger, 1, 'N/A', 'N/A'];
      __valid = false;
    } else if (!__isValidDatetime()) {
      __trigger = '時間外';
      __status = [__trigger, 1, 'N/A', 'N/A'];
      __valid = false;
    }
  }

  function __validateEdit() {
    __trigger = String(__values[settings.config.colStatus]);
    const validTriggers = Object.keys(settings.templates);

    if (__finalizeTriggers.includes(__trigger)) {
      // 予約確定のトリガーなら
      // リマインダーの設定
      const remindDate = new Date(__from);
      remindDate.setDate(__from.getDate() - 1);
      const today = new Date();
      today.setHours(19);
      __status = [1, remindDate, '送信準備'];
      if (remindDate <= today) {
        // リマインド日が，予約確定させた日の19時よりも前の場合
        __status[2] = '直前のため省略';
      }
      __valid = true;
    } else if (validTriggers.includes(__trigger)) {
      // 予約確定トリガーではないが，有効なトリガーの場合
      __status = [1, 'N/A', 'N/A'];
      __valid = false; // トリガーはvalidだが，実験の応募はvalidではない
    } else {
      // 登録されたトリガーではない場合
      throw new Error('予約ステータスに入力された文字列（トリガー）が「テンプレート」に存在しないため，メールの送信等の処理は行われませんでした。');
    }
  }

  function __allocateOnSubmission(numRow) {
    sheets.sheets[0].getRange(numRow, settings.config.colStatus + 1, 1, __status.length).setValues([__status]);
    // カレンダーの編集
    if (__valid) {
      const eventTitle = '仮予約: ' + __name;
      __calendar.createEvent(eventTitle, __from, __to);
    }
  }

  function __allocateOnEdit(numRow) {
    sheets.sheets[0].getRange(numRow, settings.config.colMailed + 1, 1, __status.length).setValues([__status]);
    // カレンダーの編集
    // まず予約イベントを削除する
    const events = __calendar.getEvents(__from, __to);
    events.forEach((e) => {
      if (e.getTitle().includes(__name)) e.deleteEvent();
    });
    if (__valid) {
      //予約確定情報をカレンダーに追加
      let newEventName = '予約完了:' + __name;
      if (settings.config.colParNameKana > 0) {
        newEventName = newEventName + '(' + __values[settings.config.colParNameKana] + ')';
      }
      __calendar.createEvent(newEventName, __from, __to);
    }
  }

  function __allocateOnTime(numRow) {
    sheets.sheets[0].getRange(numRow, settings.config.colReminded + 1).setValue('送信済み'); // シートの修正
  }

  function __fmtExpDateTimeType1() {
    const date = __values[settings.config.colExpDate];
    __from = new Date(date);
    __to = new Date(__from);
    __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
  }

  function __fmtExpDateTimeType2() {
    const date = zenToHan(__values[settings.config.colExpDate]);
    const time = zenToHan(__values[settings.config.colExpTime]);
    // 日付の処理
    __from = new Date();
    const date_info = date.match(/\d+/g); // 数字の部分だけを取り出す
    if (date_info.length == 3) {
      const [year, month, day] = date_info;
      __from.setFullYear(year, month - 1, day);
    } else if (date_info.length == 2) {
      const [month, day] = date_info;
      __from.setMonth(month - 1, day);
    } else if (date_info.length == 1) {
      const [day] = date_info;
      __from.setDate(day);
    }

    // 時間の処理
    const from_to = time.match(/\d+/g); // 数字の部分だけを取り出す
    if (from_to.length == 4) {
      // timeが hh:mm-hh:mm 形式なら
      const [fromHour, fromMin, toHour, toMin] = from_to;
      __from.setHours(fromHour, fromMin);
      __to.setHours(toHour, toMin);
    } else if (from_to.length == 2) {
      // timeが hh:mm 形式なら
      const [fromHour, fromMin] = from_to;
      __from.setHours(fromHour, fromMin);
      __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
    }
  }
  /*
    yyyy-MM-dd HH:mm -> 5
    yyyy/MM/dd HH:mm
    yyyy/MM/dd HH時mm分
    yyyy年MM月dd日 HH時mm分
    yyyy年MM月dd日HH時mm分

    MM/dd HH:mm -> 4
    MM月dd日HH時mm分 -> 4

    yyyy/MM/dd HH:mm-HH:mm -> 7
    yyyy年MM月dd日HH時mm分-HH時mm分
    
    HH:mm-HH:mm -> 4
  */
  function __fmtExpDateTimeType3() {
    const datetime = zenToHan(__values[settings.config.colExpDate]);
    __from = new Date();
    const from_to = datetime.match(/\d+/g); // 数字の部分だけを取り出す
    if (from_to.length == 7) {
      // timeが yyyy/MM/dd HH:mm-HH:mm 形式なら
      const [year, month, day, fromHour, fromMin, toHour, toMin] = from_to;
      __from.setFullYear(year, month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setHours(toHour, toMin);
    } else if (from_to.length == 6) {
      // timeが MM/dd HH:mm-HH:mm 形式なら
      const [month, day, fromHour, fromMin, toHour, toMin] = from_to;
      __from.setFullYear(month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setHours(toHour, toMin);
    } else if (from_to.length == 5) {
      // timeが yyyy-MM-dd HH:mm 形式なら
      const [year, month, day, fromHour, fromMin] = from_to;
      __from.setFullYear(year, month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
    } else if (from_to.length == 4) {
      // timeが MM-dd HH:mm 形式なら
      const [month, day, fromHour, fromMin] = from_to;
      __from.setFullYear(month - 1, day);
      __from.setHours(fromHour, fromMin);
      __to = new Date(__from);
      __to.setMinutes(__from.getMinutes() + settings.config.experimentLength);
    }
  }

  return {
    get name() {
      return __name;
    },
    get address() {
      return __address;
    },
    get from() {
      return __from;
    },
    get to() {
      return __to;
    },
    set values(val) {
      __values = val;
      __name = __values[settings.config.colParName];
      __address = __values[settings.config.colAddress];
      if (TYPE == 1) {
        __fmtExpDateTimeType1();
      } else if (TYPE == 2) {
        __fmtExpDateTimeType2();
      } else if (TYPE == 3) {
        __fmtExpDateTimeType3();
      }
      if (__from !== undefined) {
        __from.setSeconds(0, 0);
        __to.setSeconds(0, 0);
      }
    },
    get values() {
      return __values;
    },
    get trigger() {
      return __trigger;
    },
    get status() {
      return __status;
    },
    get assistant() {
      return __values[settings.config.colAssistant];
    },

    setEventType: function (eventType) {
      __event_type = eventType;
      return this;
    },

    validate: function () {
      if (__event_type == ScriptApp.EventType.ON_FORM_SUBMIT) {
        __validateSubmission();
      } else if (__event_type == ScriptApp.EventType.ON_EDIT) {
        __validateEdit();
      }

      return this;
    },

    allocate: function (numRow) {
      if (__event_type == ScriptApp.EventType.ON_FORM_SUBMIT) {
        __allocateOnSubmission(numRow);
      } else if (__event_type == ScriptApp.EventType.ON_EDIT) {
        __allocateOnEdit(numRow);
      } else if (__event_type == ScriptApp.EventType.CLOCK) {
        __allocateOnTime(numRow);
      }

      return this;
    },
  };
})();

// スクリプトトリガーをいじるオブジェクト
const scriptTriggers = (function () {
  let __triggers;

  return {
    get triggers() {
      if (__triggers === undefined) {
        __triggers = ScriptApp.getProjectTriggers();
      }
      return __triggers;
    },

    init: function () {
      this.triggers.forEach((tr) => ScriptApp.deleteTrigger(tr)); // 削除する

      // 新しく設定する
      ScriptApp.newTrigger('onOpening').forSpreadsheet(sheets.ss).onOpen().create();
      ScriptApp.newTrigger('onFormSubmission').forSpreadsheet(sheets.ss).onFormSubmit().create();
      ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(sheets.ss).onEdit().create();
      ScriptApp.newTrigger('onClock').timeBased().atHour(19).nearMinute(30).everyDays(1).inTimezone('Asia/Tokyo').create();
      // ScriptApp.newTrigger('onCalendarUpdated').forUserCalendar(Session.getActiveUser().getEmail()).onEventUpdated().create();
    },

    updateClockTrigger: function (newHour, timeZone) {
      this.triggers.forEach((tr) => {
        if (tr.getEventType() == ScriptApp.EventType.CLOCK) {
          ScriptApp.deleteTrigger(tr);
          ScriptApp.newTrigger('onClock').timeBased().atHour(newHour).nearMinute(30).everyDays(1).inTimezone(timeZone).create();
        }
      });
    },

    // updateCalendarTrigger: function (calendar_id) {
    //   this.triggers.forEach((tr) => {
    //     if (tr.getEventType() == ScriptApp.EventType.ON_EVENT_UPDATED) {
    //       ScriptApp.deleteTrigger(tr);
    //       ScriptApp.newTrigger('onCalendarUpdated').forUserCalendar(calendar_id).onEventUpdated().create();
    //     }
    //   });
    // },
  };
})();

const dlg = (function () {
  let __ui;

  return {
    get ui() {
      if (__ui === undefined) {
        __ui = SpreadsheetApp.getUi();
      }
      return __ui;
    },

    alert: function (title, prompt, buttons) {
      return this.ui.alert(title, prompt, buttons);
    },
  };
})();

const schedule = (function () {
  let __calendar;
  let __available;

  function __getAvailable(available_array) {
    __available = new Map();
    for (let row = 0; row < available_array.length; row++) {
      let available_date = available_array[row][0];
      // A列に日付が適切に入力されていないなら処理をスキップする
      if (!is(available_date, 'Date')) {
        continue;
      }
      let key = fmtDate(available_date, 'yyyy/MM/dd');
      let datetimes = [];
      for (let col = 1; col < available_array[row].length; col++) {
        let available_time = available_array[row][col];
        if (is(available_time, 'Date')) {
          available_time.setFullYear(available_date.getFullYear(), available_date.getMonth(), available_date.getDate());
          // 空き予定でない場合は空文字にする
          if (!__isAvailable(available_time)) {
            available_time = '';
          }
        }
        datetimes.push(available_time);
      }
      __available.set(key, datetimes);
    }
  }

  function __isAvailable(datetime) {
    if (__calendar === undefined) {
      __calendar = CalendarApp.getCalendarById(settings.config.workingCalendar);
    }
    const stime = new Date(datetime);
    const etime = new Date(stime);
    etime.setMinutes(etime.getMinutes() + settings.config.experimentLength);
    settings.config.closeTime.setFullYear(etime.getFullYear(), etime.getMonth(), etime.getDate());

    // 計算された終了時刻が設定されている終了時刻を超えていないか
    if (etime > settings.config.closeTime) {
      return false;
    }

    // カレンダーの予定と重複しているかどうか
    const events = __calendar.getEvents(stime, etime);
    if (events.length == 0) {
      return true;
    }
    return false;
  }

  return {
    get available() {
      if (__available === undefined) {
        __getAvailable(sheets.getValuesOf('空き予定', true));
      }
      return __available;
    },

    set calendar(val) {
      __calendar = val;
    },

    update: function () {
      __getAvailable(sheets.getValuesOf('空き予定', true));
      return this;
    },

    allocate: function () {
      const table = [];
      for (let [exp_date, exp_times] of __available.entries()) {
        let new_row = exp_times.map((exp_time) => {
          if (is(exp_time, 'Date')) {
            return fmtDate(new Date(exp_time), 'HH:mm');
          }
          return exp_time; // should be blank string
        });
        new_row.splice(0, 0, exp_date);
        table.push(new_row);
      }
      const sh = sheets.getSheetByName('空き予定');
      sh.getRange(1, 1, table.length, table[0].length).setValues(table);
    },

    init: function () {
      const available_array = [];
      for (let now = new Date(settings.config.openDate); now <= settings.config.closeDate; now.setDate(now.getDate() + 1)) {
        const new_row = [];
        new_row.push(new Date(now));
        settings.config.closeTime.setFullYear(now.getFullYear(), now.getMonth(), now.getDate());
        while (now < settings.config.closeTime) {
          new_row.push(new Date(now));
          now.setMinutes(now.getMinutes() + settings.config.experimentLength);
        }
        now.setHours(settings.config.openTime.getHours(), settings.config.openTime.getMinutes(), 0, 0); // for 文の終了条件での比較のため
        available_array.push(new_row);
      }
      __getAvailable(available_array);
      sheets.getSheetByName('空き予定').clearContents();
      this.allocate();
    },
  };
})();

///////////////////////////////////////////////////////////////////////////////
// 初期設定に関わる関数
///////////////////////////////////////////////////////////////////////////////

// 設定用のシートおよびその見本を最初に作る関数
settings.init = function () {
  try {
    let buttons = dlg.ui.ButtonSet.OK_CANCEL;
    let start = true;
    let msg;

    // タイプの確認
    if (TYPE == 1) {
      msg = '自由回答形式の設定で初期化を行います';
    } else if (TYPE == 2) {
      msg = '選択形式の設定で初期化を行います';
    } else if (TYPE == 3) {
      msg = '選択形式の設定で初期化を行います';
    } else {
      msg = '半角数字の1,2,3のいずれかを入力して設定の形式を選択してください';
      buttons = dlg.ui.ButtonSet.OK;
      start = false;
    }
    let choice = dlg.alert('設定の初期化', msg, buttons);
    // let choice = Browser.msgBox('設定の初期化', msg, buttons);
    if (choice != dlg.ui.Button.OK) {
      start = false;
    }

    if (sheets.length > 1 && start) {
      msg = '一度設定を行ったことがあるようです（シートが2枚以上あります）。\nもう一度初期化を行いますか？\n';
      msg += 'フォームの回答が一番初めのシートでないとこれまでの情報が失われる場合があります。';
      choice = dlg.alert('設定の初期化を行います', msg, buttons);
      // let choice = Browser.msgBox('設定の初期化を行います', msg, buttons);
      if (choice != dlg.ui.Button.OK) {
        start = false;
      }
    }

    if (start) {
      sheets.ss.setSpreadsheetTimeZone('Asia/Tokyo');
      settings.default.create();
      scriptTriggers.init();
      msg = '初期設定が終了しました。\n';
      msg += '「設定」シートの太枠に囲まれた項目を適切な情報に変更してください。';
      dlg.alert('設定の初期化', msg, dlg.ui.ButtonSet.OK);
      // Browser.msgBox('設定の初期化', msg, Browser.Buttons.OK);
    } else {
      dlg.alert('設定の初期化', '初期化はキャンセルされました', dlg.ui.ButtonSet.OK);
      // Browser.msgBox('設定の初期化', '初期化はキャンセルされました', Browser.Buttons.OK);
    }
  } catch (err) {
    //実行に失敗した時に通知
    const msg = `[${err.name}] ${err.stack}`;
    console.error(msg);
    dlg.alert('エラーが発生しました', msg, dlg.ui.ButtonSet.OK);
    // Browser.msgBox('エラーが発生しました', msg, Browser.Buttons.OK);
  }
};

settings.default = (function () {
  const __sheet_answers = sheets.sheets[0];
  const __default = {};

  function __createDefault() {
    const default_timezone = 'Asia/Tokyo';
    const close_date = new Date();
    close_date.setDate(new Date().getDate() + 13);
    __default.config = [
      ['設定項目', 'メール本文内でのキー', '値'],
      ['実験責任者名', 'experimenterName', '実験太郎'],
      ['実験責任者のGmailアドレス', 'experimenterMailAddress', Session.getActiveUser().getEmail()],
      ['実験責任者の電話番号', 'experimenterPhone', 'xxx-xxx-xxx'],
      ['実験の実施場所', 'experimentRoom', '実施場所'],
      ['実験の所要時間', 'experimentLength', 60],
      ['実験開始可能時刻', 'openTime', 9],
      ['実験終了時刻', 'closeTime', 19],
      ['参照するカレンダー', 'workingCalendar', Session.getActiveUser().getEmail()],
      ['実験開始日', 'openDate', Utilities.formatDate(new Date(), default_timezone, 'yyyy/MM/dd')],
      ['実験最終日', 'closeDate', Utilities.formatDate(close_date, default_timezone, 'yyyy/MM/dd')],
      ['リマインダー送信時刻', 'remindHour', 19],
      ['予約を完了させるトリガー', 'finalizeTrigger', 111],
      ['タイムゾーン設定', 'expTimeZone', 'Asia/Tokyo'],
      ['自動送信メール残数', 'remainingMails', MailApp.getRemainingDailyQuota()],
      ['予約確認メールを自分にも送るか', 'selfBccTentative', 1],
      ['予約完了メールを自分にも送るか', 'selfBccFinalize', 0],
      ['リマインダーを自分にも送るか', 'selfBccReminder', 0],
      ['翌日の実験予定を送るか', 'sendTmrwExps', 1],
      ['フォーム周りの関数を使用するか', 'useFormSystem', 1],
      ['参加者名の列', 'colParName', 'B'],
      ['ふりがなの列', 'colParNameKana', null],
    ];

    __default.config_note_template = '「フォームの回答」シートにある該当の列と一致しているか確認してください';
    __default.config_notes = [
      ['各項目の備考がコメントとして付されています'],
      ['実験責任者の名前を記入してください'], // 実験責任者
      ['変更する必要はありません。実験用のGmailアドレスが入力されています'], // 実験責任者のGmailアドレス
      ['電話番号を記入してください'], // 電話番号
      ['実験の実施場所を記入してください'], // 実施場所
      ['実験の所要時間を記入してください。'], // 実験の所要時間
      ['何時から実験できるかを記入してください（24時間表記）'], // 実験開始時刻
      ['何時まで実験可能かを記入してください（24時間表記）'], // 実験終了時刻
      ['利用したいカレンダーのIDをコピペしてください'], // 参照するカレンダー
      ['実験を開始する日付を記入してください（年/月/日で表記）'], // 実験開始日
      ['実験の終了予定日を記入してください（年/月/日で表記）'], // 実験最終日
      ['リマインダーを送信する時刻を記入してください（24時間表記）。実験終了時刻以後にして下さい。なお指定した時刻から1時間以内に送信されます。'], // リマインダー送信時刻
      ['必要に応じて任意の半角数字列に変更してください。複数指定する場合はカンマで区切ってください。'], // 予約を完了させるトリガー
      ['必要に応じて変更してください。形式は http://joda-time.sourceforge.net/timezones.html を参照してください。'], // タイムゾーン設定
      ['自動で送信できるメールの残数の目安です。「担当」機能を使っていると一気に2減ったりします。1日経つと100に近い値に戻ります。'], // 自動送信メール残数
      ['自分にも予約確認メールを送る場合は1を，送らない場合は0を入力してください。送らない場合は自動送信できる総メール数が増えます（以下同様）。'], // 予約確認メールを自分にも送るか
      ['自分にも予約完了メールを送る場合は1を，送らない場合は0を入力してください。'], // 予約完了メールを自分にも送るか
      ['自分にも参加者と同様のリマインダーを送る場合は1を，送らない場合は0を入力してください。'], // リマインダーを自分にも送るか
      ['翌日の実験予定の一覧を自分にメールする場合は1を，しない場合は0を入力してください。'], // 翌日の実験予定を送るか
      [
        'ここを0にすると，formに関わる関数が動作しなくなります。この項目はスプレッドシートだけからメールの自動送信システムだけを使用したい人を想定しています',
      ], // フォーム周りの関数を使用するか
      [__default.config_note_template], // 参加者名の列
      [__default.config_note_template + 'もし利用しない場合は空欄にしてください。'], // ふりがなの列
    ];

    // メールテンプレート
    const template_bodies = {
      仮予約: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験への応募ありがとうございました。',
        '予約の確認メールを自動で送信しております。\n',
        'expDate fromWhen〜toWhen',
        'で予約を受け付けました（まだ確定はしていません)。',
        '後日、予約完了のメールを送信いたします。',
        'もし日時の変更等がある場合は experimenterMailAddress までご連絡ください。',
        'どうぞよろしくお願いいたします。\n',
        'experimenterName',
      ],
      時間外: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験への応募ありがとうございました。',
        '申し訳ありませんが、ご希望いただいた',
        'expDate fromWhen〜toWhen',
        'は実験実施可能時間（openTime時〜closeTime時）外または、実施期間（openDate〜closeDate）外です。',
        'お手数ですが、もう一度登録し直していただきますようお願いします。\n',
        'experimenterName',
      ],
      重複: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験への応募ありがとうございました。',
        '申し訳ありませんが、ご希望いただいた',
        'expDate fromWhen〜toWhen',
        'にはすでに予約（予定）が入っており（タッチの差で他の方が予約をされた可能性もあります）、実験を実施することができません。',
        'お手数ですが、もう一度別の日時で登録し直していただきますようお願いします。\n',
        'experimenterName',
      ],
      予約完了wd: [
        'participantName 様\n',
        'この度は心理学実験への応募ありがとうございました。',
        'expDate fromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
        '場所はexperimentRoomです。当日は直接お越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '当日もよろしくお願いいたします。\n',
        '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)',
        '当日の連絡はexperimenterPhoneまでお願いいたします。',
      ],
      予約完了we: [
        'participantName 様\n',
        'この度は心理学実験への応募ありがとうございました。',
        'expDate fromWhen〜toWhenの心理学実験の予約が完了しましたのでメールいたします。',
        '場所はexperimentRoomです。休日は教育学部棟玄関の鍵がかかっており、外から入ることができません。実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '当日もよろしくお願いいたします。\n',
        '実験責任者experimenterName（当日は他の者が実験担当する可能性があります)',
        '当日の連絡はexperimenterPhoneまでお願いいたします。',
      ],
      222: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験への応募ありがとうございました。',
        '大変申し訳ありませんが、以前実施した同様の実験にご参加いただいており、今回の実験にはご参加いただけません。ご了承ください。\n',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '今後ともよろしくお願いします。\n',
        'experimenterName',
      ],
      333: [
        'participantName 様\n',
        '心理学実験実施責任者のexperimenterNameです。',
        'この度は心理学実験への応募ありがとうございました。',
        '大変申し訳ありませんが、応募いただいた段階ですでに募集人数の定員に達していたため、実験に参加していただくことができません。ご了承ください。\n',
        '今後、次の実験を実施する際に再度応募していただけると幸いです。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        '今後ともよろしくお願いいたします。\n',
        'experimenterName',
      ],
      リマインダーwd: [
        'participantName 様\n',
        '実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
        '明日 fromWhenから実験に参加していただく予定となっております。',
        '場所はexperimentRoomです。実験時間に実験室まで直接お越しください。\n',
        'なお、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        'それでは明日、よろしくお願いいたします。\n',
        'experimenterName',
      ],
      リマインダーwe: [
        'participantName 様\n',
        '実験者のexperimenterNameです。明日参加していただく実験についての確認のメールをお送りしています。\n',
        '明日 fromWhenから実験に参加していただく予定となっております。',
        '場所はexperimentRoomです。\n',
        'なお、明日は休日のため教育学部棟玄関の鍵がかかっており、外から入ることができません。実験者が実験開始5分前から玄関前で待機しておりますので、実験開始時間までにお越しください。\n',
        'また、実験中は眠くなりやすいため、本日は十分な睡眠を取って実験にお越しください。',
        'ご不明な点などありましたら、experimenterMailAddressまでご連絡ください。',
        'それでは明日、よろしくお願いいたします。\n',
        'experimenterName',
      ],
    };

    for (const key in template_bodies) {
      template_bodies[key] = template_bodies[key].join('\n');
    }

    const not_used = '利用する場合はここに本文を記載するとともに土日での変更の数字を1に変えてください。なお，改行は"alt + enter"です';

    __default.templates = [
      ['トリガー', '休日での変更', '題名', '本文（平日）', '本文（土日祝）'],
      ['仮予約', 0, '予約の確認', template_bodies['仮予約'], not_used],
      ['時間外', 0, '実験実施可能時間外です', template_bodies['時間外'], not_used],
      ['重複', 0, '予約が重複しています', template_bodies['重複'], not_used],
      [111, 1, '実験予約が完了いたしました', template_bodies['予約完了wd'], template_bodies['予約完了we']],
      [222, 0, '以前に実験にご参加いただいたことがあります', template_bodies[222], not_used],
      [333, 0, '定員に達してしまいました', template_bodies[333], not_used],
      ['リマインダー', 1, '明日実施の心理学実験のリマインダー', template_bodies['リマインダーwd'], template_bodies['リマインダーwe']],
    ];

    const note =
      '適宜変更してください。参加者名は participantName ，実験実施時間は fromWhen および toWhen に代入されます。その他のキーは設定シートを参照してください。';

    __default.templates_note = __default.templates.map((_, idx) => {
      if (idx == 0) {
        return [null, null];
      }
      return [note, note];
    });

    // メンバー
    const sh_name_answers = __sheet_answers.getName();
    const last_column = __sheet_answers.getLastColumn();
    const last_col_notation = __sheet_answers.getRange(1, last_column).getA1Notation().replace(/\d/, ''); // 列のアルファベットを取得
    const formula = `=COUNTIF('${sh_name_answers}'!${last_col_notation}:${last_col_notation}, A2)`;
    __default.members = [
      ['キー', '名前', 'アドレス', '担当回数'],
      [1, 'りんご', 'apple@hogege.com', formula],
      [2, 'ごりら', 'gorilla@hogege.com', ''],
      [3, 'らっぱ', 'horn@hogege.com', ''],
    ];

    // 空き予定
    __default.available = [];
    for (const now = new Date(); now <= close_date; now.setDate(now.getDate() + 1)) {
      const new_row = [];
      new_row.push(Utilities.formatDate(now, default_timezone, 'yyyy/MM/dd'));
      now.setHours(9, 0, 0);
      const close_time = 19;
      const exp_length = 60;
      for (const cur_time = new Date(now); cur_time.getHours() < close_time; cur_time.setMinutes(cur_time.getMinutes() + exp_length)) {
        new_row.push(Utilities.formatDate(now, default_timezone, 'HH:mm'));
      }
      __default.available.push(new_row);
    }
  }

  function __addNewColNames() {
    const current_colnms = __sheet_answers.getRange(1, 1, 1, __sheet_answers.getLastColumn()).getValues()[0];
    const new_colnms = ['予約ステータス', '連絡したか', 'リマインド日時', 'リマインドしたか', '担当'];
    const colnms = current_colnms.concat(new_colnms);
    __sheet_answers.getRange(1, 1, 1, colnms.length).setValues([colnms]);
  }

  function __createConfig() {
    sheets.ss.insertSheet('設定');
    const sh = sheets.ss.getSheetByName('設定');
    const last_column = __sheet_answers.getLastColumn() - 1;
    const extra_config_type1 = [
      ['参加者アドレスの列', 'colAddress', numToColumnNotation(last_column - 6)],
      ['希望日時の列', 'colExpDate', numToColumnNotation(last_column - 5)],
    ];

    const extra_config_type2 = [
      ['参加者アドレスの列', 'colAddress', numToColumnNotation(last_column - 7)],
      ['希望日の列', 'colExpDate', numToColumnNotation(last_column - 6)],
      ['希望時間の列', 'colExpTime', numToColumnNotation(last_column - 5)],
    ];

    const extra_config_common = [
      ['予約ステータスの列', 'colStatus', numToColumnNotation(last_column - 4)],
      ['「連絡したか」の列', 'colMailed', numToColumnNotation(last_column - 3)],
      ['リマインド日時の列', 'colRemindDate', numToColumnNotation(last_column - 2)],
      ['「リマインドしたか」の列', 'colReminded', numToColumnNotation(last_column - 1)],
      ['担当の列', 'colAssistant', numToColumnNotation(last_column)],
    ];

    let extra_config;
    if (TYPE == 1 || TYPE == 3) {
      extra_config = extra_config_type1.concat(extra_config_common);
    } else if (TYPE == 2) {
      extra_config = extra_config_type2.concat(extra_config_common);
    }

    __default.config.push(...extra_config);
    const extra_config_notes = extra_config.map(() => [__default.config_note_template]);
    __default.config_notes.push(...extra_config_notes);

    // 値の設定
    const nrow = __default.config.length;
    const ncol = __default.config[0].length;
    sh.getRange(1, 1, nrow, ncol).setValues(__default.config);

    // 注釈
    sh.getRange(1, 3, nrow, 1).setNotes(__default.config_notes);

    // 書式の設定
    sh.getRange(15, 3).setFontColor('#FF0000'); // メールの残数のセルを赤色にする
    sh.getRange(2, 2, nrow - 1, 1).setFontColor('#C8C8C8');
    sh.autoResizeColumn(1);
    sh.autoResizeColumn(3);
    sh.getRange(2, 3, 1, 1).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(4, 3, 8, 1).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(16, 3, 5, 1).setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  function __createMailTemplate() {
    sheets.ss.insertSheet('テンプレート');
    const sh = sheets.ss.getSheetByName('テンプレート');
    const rng = sh.getRange(1, 1, __default.templates.length, __default.templates[0].length);
    rng.setValues(__default.templates);
    // 体裁を整える
    rng.setVerticalAlignment('top');
    sh.setColumnWidth(4, 500);
    sh.setColumnWidth(5, 500);
    const cell_text_wrap = __default.templates.map(() => [false, false, true, true, true]);
    rng.setWraps(cell_text_wrap);

    // 注釈の設定
    sh.getRange(1, 4, __default.templates_note.length, __default.templates_note[0].length).setNotes(__default.templates_note);
  }

  function __createMembers() {
    // メンバーシートの設定
    sheets.ss.insertSheet('メンバー');
    const sh = sheets.ss.getSheetByName('メンバー');
    sh.getRange(1, 1, __default.members.length, __default.members[0].length).setValues(__default.members);
    sh.getRange(1, 3).setNote('Gmailのアドレスでなくても大丈夫です。');
  }

  function __createAvailable() {
    // 空き予定シートの設定
    sheets.ss.insertSheet('空き予定');
    const sh = sheets.ss.getSheetByName('空き予定');
    sh.getRange(1, 1, __default.available.length, __default.available[0].length).setValues(__default.available);
  }

  return {
    create() {
      if (sheets.length > 2) {
        sheets.sheets.forEach((sh, idx) => {
          if (idx > 0) {
            sheets.ss.deleteSheet(sh);
          }
        });
      } else {
        // フォームの回答に新しい列を追加する
        __addNewColNames();
      }
      __createDefault();
      __createConfig();
      __createMailTemplate();
      __createMembers();
      if (TYPE == 3) {
        __createAvailable();
        SpreadsheetApp.getUi().createMenu('カレンダー').addItem('カレンダーをシートに反映', 'onCalendarUpdated').addToUi();
      }
      sheets.ss.insertSheet('Cached');
      sheets.update();
      settings.retrieve().save();
      sheets.ss.getSheetByName('設定').activate(); // 設定画面を開く
    },
  };
})();
