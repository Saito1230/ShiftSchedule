let ONE_OPERATION_COLOR = "#FFDFDF";
let FILL_COLOR = "#5B9BD5";
let WEEK_FILL_COLOR = "#FFDFDF";
let PRIORITY_FILL_COLOR = "#F0E4FC";

const InserSheetName = "シフト表";
const TableSheetName = "シート1";
const ScriptSheetName = "スクリプト";

const TimeList = ["9:00", "9:10", "9:20", "9:30", "9:40", "9:50", "10:00", "10:10", "10:20", "10:30", "10:40", "10:50", "11:00", "11:10", "11:20", "11:30", "11:40", "11:50", "12:00", "12:10", "12:20", "12:30", "12:40", "12:50", "13:00", "13:10", "13:20", "13:30", "13:40", "13:50", "14:00", "14:10", "14:20", "14:30", "14:40", "14:50", "15:00", "15:10", "15:20", "15:30", "15:40", "15:50", "16:00", "16:10", "16:20", "16:30", "16:40", "16:50", "17:00", "17:10", "17:20", "17:30", "17:40", "17:50", "18:00", "18:10", "18:20", "18:30", "18:40", "18:50", "19:00"];

const DayWeek = ["mon", "tue", "wed", "thu", "fri"];
const KanjiWeek = {"mon":"月", "tue": "火", "wed": "水", "thu": "木", "fri":"金"}

let ss = SpreadsheetApp.getActiveSpreadsheet();
let sheet = ss.getActiveSheet();
let TableSheet = ss.getSheetByName(TableSheetName);

function myFunction() {
  if(Browser.msgBox("実行しますか", Browser.Buttons.YES_NO)=="no") return;
  Init();
  let data = getData() ?? false;

  //err
  if(!data) return;

  DrawChart(data);
}

function getData(){
  //skip explain -> 2;
  let index = 2;
  let res = [];

  while(true){
    let s_name = TableSheet.getRange(index, 3).getDisplayValue();
    let s_mon = TableSheet.getRange(index, 6).getDisplayValue().split(",");
    let s_tue = TableSheet.getRange(index, 7).getDisplayValue().split(",");
    let s_wed = TableSheet.getRange(index, 8).getDisplayValue().split(",");
    let s_thu = TableSheet.getRange(index, 9).getDisplayValue().split(",");
    let s_fri = TableSheet.getRange(index, 10).getDisplayValue().split(",");

    if(s_name == "") break;

    let d = [s_mon, s_tue, s_wed, s_thu, s_fri];
    let date = {};

    //convert format
    //{
    //  "mon": [ {"start": "xx:xx", "end": "xx:xx"}, {"start": "xx:xx", "end": "xx:xx"} ],
    //}
    for(let i = 0; i < DayWeek.length; i++){
      let t = String(d[i]).replace("、", ",");
      let s = t.split(",");

      if(s == "なし") {
        date[DayWeek[i]] = null;
      }else{
        let add = [];
        for(let q = 0; q < s.length; q++){

          let item = null;
          if(~s[q].indexOf("~")) item = s[q].split("~");
          else if(~s[q].indexOf("～")) item = s[q].split("～")
          else item = s[q].split("-");

          let f = false;
          //◎, (終日) check
          if((item[1] ?? false) && item[1].includes("◎")) {
            item[1] = item[1].replace("◎", "");
            f = true;
          }else if((item[1] ?? false) && item[1].includes("終日")) {
            item[1] = item[1].replace("(終日)", "");
          }

          //remove space
          item[0] = item[0].replace(" ", "").replace("　", "").replace("○", "").replace("〇", "");
          item[1] = item[1].replace(" ", "").replace("　", "").replace("○", "").replace("〇", "");

          //09:00を9:00に変換
          if(item[0] == "09:00") item[0] = "9:00"
          else if(item[1] == "09:00") item[1] = "9:00";

          //calc time index
          let start_index = TimeList.indexOf(item[0]);
          let end_index = TimeList.indexOf(item[1]);

          //5分単位の場合は5分前の時間で処理
          start_index = (start_index < 0) ?
              TimeList.indexOf(item[0].substring(0, item[0].lastIndexOf("5")) + "0") :
              start_index;

          end_index = (end_index < 0) ?
              TimeList.indexOf(item[1].substring(0, item[1].lastIndexOf("5")) + "0") :
              end_index;


          //時間範囲チェック
          if(start_index < 0 || end_index < 0){
            let msg = `${s_name}さんの${KanjiWeek[DayWeek[i]]}曜日のシフトは時間範囲外です。\r\n ${TimeList[0]}～${TimeList[TimeList.length-1]}の範囲で設定してください。`;
            if(Browser.msgBox(msg, Browser.Buttons.OK)=="ok") return null;
          }

          let active_index = [];
          for(let p = start_index; p < end_index; p++) active_index.push(p);

          add.push(
              {
                "start": item[0],
                "end": item[1],
                "flag": f,
                "start_index": start_index,
                "end_index": end_index,
                "active_index": active_index
              }
          );

        }
        date[DayWeek[i]] = add;
      }
    }//convert format END

    res.push(new Task(s_name, date));
    index++;

  }//while END

  return res;
}

function DrawChart(data){
  let ChartSheet = ss.getSheetByName(InserSheetName);

  //add colom
  ChartSheet.insertColumnsAfter(26, 50);

  //set width
  ChartSheet.setColumnWidths(3, 4, 30);
  ChartSheet.setColumnWidths(5, 5, 60);
  ChartSheet.setColumnWidths(6, 9, 60);
  ChartSheet.setColumnWidths(10, 60, 8);


  //write member info
  let table_space_x = 2;
  for(let week_index = 0; week_index < DayWeek.length; week_index++){

    ss.toast(`${StrProgress(week_index, DayWeek.length-1)}`, "進捗");
    let week = DayWeek[week_index];

    /** calc one operation **/
    let sum_active_index = [];
    let sum_member = 0;
    for(let i = 0; i < data.length; i++){
      if(data[i].attend[week] === null) continue;
      sum_member++;
      for(let q = 0; q < data[i].attend[week].length; q++){
        for(let w = 0; w < data[i].attend[week][q].active_index.length; w++){
          sum_active_index.push(data[i].attend[week][q].active_index[w]);
        }
      }
    }

    //write text
    ChartSheet.getRange(table_space_x, 5).setValue("名前");
    ChartSheet.getRange(table_space_x, 6).setValue("時");

    //merge cell & write hour
    let h = 9;
    for(let i = 10; i < 67; i+=6){
      ChartSheet.getRange(table_space_x, i, 1, 6).merge().setValue(h).setBorder(null, null, null, true, null,null);
      h++;
    }

    let dict = {};
    for(let key of sum_active_index){
      dict[key] = sum_active_index.filter((x) => {
        return x == key
      }).length;
    }

    let keys = Object.keys(dict);
    for(let i = 0; i < keys.length; i++){
      if(dict[keys[i]] == 1){
        //get range(x,y,width)
        ChartSheet.getRange(table_space_x+1, Number(keys[i])+10, sum_member).setBackground(ONE_OPERATION_COLOR);
      };
    }
    /** calc oneoparation END */

    let num_index = 0;

    //fill
    for(let i = 0; i < data.length; i++){

      ChartSheet.getRange(table_space_x,3).setValue(KanjiWeek[week]).setBackground(WEEK_FILL_COLOR);

      //ignore not have tasks.
      if(data[i].attend[week] === null) continue;

      //write num
      ChartSheet.getRange(table_space_x+num_index+1, 4).setValue(num_index+1);

      //write name
      ChartSheet.getRange(table_space_x+num_index+1, 5).setValue(data[i].name);

      //write time
      let space = 0;
      for(let q = 0; q < data[i].attend[week].length; q++){

        //write time
        ChartSheet.getRange(table_space_x+num_index+1,6+space).setValue(data[i].attend[week][q].start);
        ChartSheet.getRange(table_space_x+num_index+1,7+space).setValue(data[i].attend[week][q].end);
        space+=2;

        //fill
        for(let w = 0; w < data[i].attend[week][q].active_index.length; w++){
          let c = (data[i].attend[week][q].flag) ? PRIORITY_FILL_COLOR : FILL_COLOR;
          ChartSheet.getRange(table_space_x+num_index+1, data[i].attend[week][q].active_index[w]+10).setBackground(c);
        }
      }
      num_index++;
    }

    ChartSheet.getRange(table_space_x, 4, sum_member+1, 66)
        .setBorder(true, true, true, true,false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ChartSheet.getRange(table_space_x, 4, sum_member+1, 1)
        .setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ChartSheet.getRange(table_space_x, 27, sum_member+1, 1)
        .setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ChartSheet.getRange(table_space_x, 45, sum_member+1, 1)
        .setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ChartSheet.getRange(table_space_x, 63, sum_member+1, 1)
        .setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ChartSheet.getRange(table_space_x, 4, 1, 66)
        .setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    for(let i = 0; i < 5; i++){
      ChartSheet.getRange(table_space_x, 5+i, sum_member+1, 1)
          .setBorder(null, null, null, true, null, null);
    }
    for(let i = 12; i < 70; i+=3){
      if(i==27 || i==45 || i==63) continue;
      ChartSheet.getRange(table_space_x, i, sum_member+1, 1)
          .setBorder(null, null, null, true, null, null);
    }


    table_space_x+=sum_member+2;
  }

  //set text center
  ChartSheet.getRange("A1:BX1000").setHorizontalAlignment("center")
}

function Init(){
  //create new sheet
  let sheeName = ss.getSheets().find((x) => {return x.getName() === InserSheetName});
  if(sheeName != null) ss.deleteSheet(ss.getSheetByName(InserSheetName));
  ss.insertSheet().setName(InserSheetName);

  /*
  let ScriptSheet = ss.getSheetByName(ScriptSheetName);
  FILL_COLOR = ScriptSheet.getRange("E5").getBackground();
  ONE_OPERATION_COLOR = ScriptSheet.getRange("E6").getBackground();
  WEEK_FILL_COLOR = ScriptSheet.getRange("E7").getBackground();
  */
}

function StrProgress(current, max){
  let per = current/max;
  let l = 15;
  let progress = Number(per*l);
  let bar = `[${"■".repeat(progress)} ${"□".repeat(l-progress)}]`;
  let percen = Number(per*100);
  return `${bar} ${percen}%`;
}

class Task{
  constructor(name, attend){
    this.name = name;
    this.attend = attend;
  }
}