// 完成版
function componetscalender() {
  // カレンダーを保存する列を指定する
    sheetkey ={introduction:"M",relation:"Q",revue:"AP",sett:"AX",patient:"BB"}
  
    const sheet = SpreadsheetApp.getActiveSheet();                  //シートをアクティブにする                                        
    const lastRow = sheet.getLastRow();                             //最終行を取得
    const contents = sheet.getRange(`A4:BE${lastRow}`).getValues()  //範囲を指定し、データを取得
    var introductionEvent = ""                                      //イベントID格納用
    mine = mineAcount()                                             //招集メンバー用
    member = defaultCalendermember()                                //デフォルトのカレンダー
    let calendar = CalendarApp.getCalendarById(member)              //カレンダーの編集用
  
    for(let i = 0; i<contents.length;i++){                            //変数に格納したデータをループし取得  
        // シート全体の変数
      const clinicName = contents[i][0]
      const salesPerson = contents[i][5]
      const sheetUrl = contents[i][1]
      
      // introductionプロジェクトの変数
      const introductionCheck = contents[i][12]
      const introductionDay = contents[i][13]
      const introductionStartTime = contents[i][14]
      const introductionEndTime = contents[i][15]
  
      //　relation プロジェクトの変数
      const relationCheck = contents[i][16]
      const relationDay = contents[i][17]
      const relationStartTime = contents[i][18]
      const relationEndTime = contents[i][19]
      
      // review　プロジェクトの変数
      const revueCheck = contents[i][41]
      const revueDay = contents[i][42]
      const revueStartTime = contents[i][43]
      const revueEndTime = contents[i][44]
  
      //setmoveの変数
      const setCheck = contents[i][49] 
      const setDay = contents[i][50]
      const setStartTime = contents[i][51]
      const setEndTime = contents[i][52]
  
      // patientの変数
      const intputPatientCheck = contents[i][53]
      const intputPatientDay = contents[i][54]
      const intputPatientStartTime = contents[i][55]
      const intputPatientEndTime = contents[i][56]
    
      // 各カレンダーのタイトル
      var introductionTitle = "【導入】"+ clinicName
      var relationTitle = "【連携】"+ clinicName 
      var revueTitle = "【レビュー会】"+ clinicName
      var setTitle = "【セット移行】"+ clinicName
      var patientTitle = "【患者取込】"+ clinicName
  
      // 営業メンバーをリストから追加
      let sales = personList(salesPerson)
      // カレンダーのデータを設定　招集メンバーが異なるため、２つに分けている
      const options ={description:`ヒアリングシート ${sheetUrl}`,guests:`${introductionMember()},${sales}`,sendInvites:false}
      const optionMine = {description:`ヒアリングシート ${sheetUrl}`,guests:`${mineAcount()},${sales}`,sendInvites:false}
      
      // introduction project
      if(introductionCheck){
      }else{
        components(introductionDay,introductionStartTime,introductionEndTime,introductionTitle,sheetkey.introduction,calendar,sheet,i,options)
      }
  
      //relation project
      if(relationCheck){
      }else{
        components(relationDay,relationStartTime,relationEndTime,relationTitle,sheetkey.relation,calendar,sheet,i,options)
      }
  
      //revue project
      if(revueCheck){    
      }else{
        components(revueDay,revueStartTime,revueEndTime,revueTitle,sheetkey.revue,calendar,sheet,i,options)
      }
  
      //setMove project
      if(setCheck){
      }else{
        components(setDay,setStartTime,setEndTime,setTitle,sheetkey.sett,calendar,sheet,i,options)
      }
  
      // patient project
      if(intputPatientCheck){
      }else{
        components(intputPatientDay,intputPatientStartTime,intputPatientEndTime,patientTitle,sheetkey.patient,calendar,sheet,i,options)
      }
    }
  }
  
  // カレンダーに登録をコンポーネント化して使い回し
  function components(day,startTime,endTime,title,action,calendar,sheet,i,options) {
    var date = new Date(day);
      if(!day){
      }else{
      if(startTime ==""|| endTime==""){ // 時間が未定の場合、終日でカレンダーを作成
        calendar.createAllDayEvent(title,date)
      }else{//日時データを関数に投げて、日本時間に修正 イベントIDを取得し、保管予定 
        const [startDateobj,endDateobj]  = timeDate(day,startTime,endTime)
        introductionEvent = calendar.createEvent(title,startDateobj,endDateobj,options).getId()
      }
        sheet.getRange(`${action}${i+4}`).setValue("TRUE"); // 作成したカレンダーをTRUE,FALSEで作成するか否かを判断
      // sheet.getRange(`BJ${i+4}`).setValue(introductionEvent); // 作成カレンダーのイベントIDの仮置き
        }
  }
  
  // 時間編集の関数
  function timeDate(day,start,end){
    // スタート時間の編集
    var startDateobj = new Date(day);
    startDateobj.setHours(start.getHours())
    startDateobj.setMinutes(start.getMinutes())
    
    // 終了時間の編集
    var endDateobj = new Date(day);
    endDateobj.setHours(end.getHours())
    endDateobj.setMinutes(end.getMinutes())
  
    return [startDateobj,endDateobj]
  
  }
  
  
  
  
  