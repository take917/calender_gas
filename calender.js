// 完成版
function newcalender() {
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
        var date = new Date(introductionDay);
        if(introductionStartTime ==""|| introductionEndTime==""){ // 時間が未定の場合、終日でカレンダーを作成
          calendar.createAllDayEvent(introductionTitle,date)
        }else{//日時データを関数に投げて、日本時間に修正 イベントIDを取得し、保管予定 
          const [startDateobj,endDateobj]  = timeDate(introductionDay,introductionStartTime,introductionEndTime)
          introductionEvent = calendar.createEvent(introductionTitle,startDateobj,endDateobj,options).getId()
        }
      sheet.getRange(`M${i+4}`).setValue("TRUE"); // 作成したカレンダーをTRUE,FALSEで作成するか否かを判断
      sheet.getRange(`BJ${i+4}`).setValue(introductionEvent); // 作成カレンダーのイベントIDの仮置き
      }
  
    //relation project
      if(relationCheck){
      }else{
        var date = new Date(relationDay);
        if(relationStartTime ==""|| relationEndTime==""){
          calendar.createAllDayEvent(relationTitle,date)
        }else{
          const [startDateobj,endDateobj]  = timeDate(relationDay,relationStartTime,relationEndTime)
          relationEvent = calendar.createEvent(relationTitle,startDateobj,endDateobj,options).getId()
        }
        sheet.getRange(`Q${i+4}`).setValue("TRUE");
      }
  
    //revue project
      if(revueCheck){    
      }else{
        var date = new Date(revueDay);
        if(revueStartTime ==""|| revueEndTime==""){
          calendar.createAllDayEvent(revueTitle,date)
        }else{
          const [startDateobj,endDateobj]  = timeDate(revueDay,revueStartTime,revueEndTime)
          revueEvent = calendar.createEvent(revueTitle,startDateobj,endDateobj,options).getId()
        }
        sheet.getRange(`AP${i+4}`).setValue("TRUE");
      }
  
    // 　setMove project
    if(setCheck){
      }else{
        var date = new Date(setDay);
        if(setStartTime ==""|| setEndTime==""){
          calendar.createAllDayEvent(setTitle,date)
        }else{
          const [startDateobj,endDateobj]  = timeDate(setDay,setStartTime,setEndTime)
          revueEvent = calendar.createEvent(setTitle,startDateobj,endDateobj,optionMine).getId()
        }
        sheet.getRange(`AX${i+4}`).setValue("TRUE");
      }
  
    // patient project
    if(intputPatientCheck){
        }else{
          var date = new Date(intputPatientDay);
          if(intputPatientStartTime ==""|| intputPatientEndTime==""){
          calendar.createAllDayEvent(setTitle,date)
        }else{
          const [startDateobj,endDateobj]  = timeDate(intputPatientDay,intputPatientStartTime,intputPatientEndTime)
          revueEvent = calendar.createEvent(patientTitle,startDateobj,endDateobj,optionMine).getId()
        }
        sheet.getRange(`BB${i+4}`).setValue("TRUE");
      }
    }
  }
  
  
  
  function timeDate(day,start,end){
        var startDateobj = new Date(day);
        startDateobj.setHours(start.getHours())
        startDateobj.setMinutes(start.getMinutes())
        
        var endDateobj = new Date(day);
        endDateobj.setHours(end.getHours())
        endDateobj.setMinutes(end.getMinutes())
  
        return [startDateobj,endDateobj]
  
  }
  
  
  
  
  