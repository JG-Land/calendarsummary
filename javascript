```jsx
// 이메일 저장용 캐시 맵
var cache = {
  emailToName: {}
}

/**
 * 유저 이메일을 이름으로 변경해줌
 * @param email 유저 이메일
 * @returns 
 */
function getUserNameByEmail(email) {
  if (email.indexOf('@brandi.co.kr') == -1) return ''
  if (cache.emailToName[email]) {
    return cache.emailToName[email];
  }
  try {
    var user = AdminDirectory.Users.get(email, {viewType:'domain_public'})
    // Logger.log('유저 이름 취득 ='+ user.name.fullName)
    cache.emailToName[email] = user.name.fullName
    return user.name.fullName
  } catch(e) {
    return ''
  }
}
/**
 * 이메일 테스트용 함수
 */
function testEmail() {
  Logger.log(getUserNameByEmail('chunbs@brandi.co.kr'))
}
/**
 * 칼럼 번호를 숫자로 바꿔줌
 * 
 * ex) 11 -> B
 * @param column 칼럼 번호 
 * @returns 
 */
function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
/**
 * 날짜를 포멧팅된 날짜로 변경해줌
 * @param date 
 * @returns 
 */
function dateFormattor(date) {
  // 대상 포멧: 12/17(목) 15:00
  var weekList = ['월','화','수','목','금','토','일']
  var week = Utilities.formatDate(date, 'Asia/Seoul', 'u')-1
  return Utilities.formatDate(date, 'Asia/Seoul', 'MM/dd')+'('+weekList[+week]+') '
  +Utilities.formatDate(date, 'Asia/Seoul', 'HH:mm')
}

 /**--------------------------------------------
 * 캔디데이트와 캘린더 동기화
 --------------------------------------------*/
 /**
  * 캔디데이츠를 캘린더와 연동
  * 
  * @param sheetName 시트명
  * @param calendarId 캘린더 아이디 (요건 캘린더애서 확인 가능함)
  * @param maxrow 최대 처리 row수 
  */
function goCandydateSyncCalendar(sheetName, calendarId, maxrow) {
  // 랩스 
  // 2->현재상태
  // 3->이름
  // 11-> 날짜
  // 12-> 면접관
  // 13-> 결과
  let startRow = 4 // 4번째 줄부터가 유효 데이터
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)  
  const data = spreadsheet.getRange('A2:AA3').getValues()
  /* 
    1,2줄 해석 (칼럼 찾기용) 
    1,2줄에서 1차면접을 찾은 후 "면접일정, 면접관, 결과, 불합격 안내" 을 찾아서 위치 저장
    2차면접, 컬처핏면접 도 동일하게 시도함
  */
  // 시트에서 데이터 존재하는 위치를 담을 변수
  const cPos = {'time1':{}, 'time2':{}, 'time3': {}}
  for (let x=0; x<data[0].length; x++) {
    if (data[0][x] == '1차면접') {
      for (xx=x; xx<data[1].length; xx++) {
        if (data[0][xx] != '1차면접' && data[0][xx] != '') break; // 1차면접 파싱끝
        if (data[1][xx] == '면접일정') cPos.time1['면접일정'] = xx
        if (data[1][xx] == '면접관') cPos.time1['면접관'] = xx
        if (data[1][xx] == '결과') cPos.time1['결과'] = xx
        if (data[1][xx] == '불합격 안내') cPos.time1['불합격 안내'] = xx
      }
    }
    if (data[0][x] == '2차면접') {
      for (let xx=x; xx<data[1].length; xx++) {
        if (data[0][xx] != '2차면접' && data[0][xx] != '') break; // 2차면접 파싱끝
        if (data[1][xx] == '면접일정') cPos.time2['면접일정'] = xx
        if (data[1][xx] == '면접관') cPos.time2['면접관'] = xx
        if (data[1][xx] == '결과') cPos.time2['결과'] = xx
        if (data[1][xx] == '불합격 안내') cPos.time2['불합격 안내'] = xx
      }
    }
    if (data[0][x] == '컬처핏면접') {
      for (let xx=x; xx<data[1].length; xx++) {
        if (data[0][xx] != '컬처핏면접' && data[0][xx] != '') break; // 2차면접 파싱끝
        if (data[1][xx] == '면접일정') cPos.time3['면접일정'] = xx
        if (data[1][xx] == '면접관') cPos.time3['면접관'] = xx
        if (data[1][xx] == '결과') cPos.time3['결과'] = xx
        if (data[1][xx] == '불합격 안내') cPos.time3['불합격 안내'] = xx
      }
    }
  }
  /* ---------------------------------------------
     캘린더 정보를 가져옴 
     대상 기간은 이번달 부터 다음달 말일까지 가져온다.
  --------------------------------------------- */
  const eventCal = CalendarApp.getCalendarById(calendarId);
  const now = new Date()
  // 총 2개월치를 가져옴
  const startDate = new Date(now.getFullYear(), now.getMonth(), 1) // 이번달 1일부터
  const endDate = new Date(now.getFullYear(), now.getMonth()+2, 1) // 다다음달 1일까지
  // 동기화 시점에 해당 달만 체크함
  const list = eventCal.getEvents(startDate, endDate)
  Logger.log('캘린더 총 '+list.length+'건 읽음')
  var calendarUserList = []
  for (i = 0; i < list.length; i++) {
    const ev = list[i]; // 캘린더 하나
    // 캘린더 내용에 [1차면접], [2차면접], [컬처핏면접] 패턴이 있으면 유효한 내용으로 분류
    if (ev.getTitle().indexOf('[1차면접]') >= 0 || ev.getTitle().indexOf('[2차면접]') >= 0 || ev.getTitle().indexOf('[컬처핏면접]') >= 0) {
      let time = '';
      if (ev.getTitle().indexOf('[1차면접]') >= 0) time = '1차';
      if (ev.getTitle().indexOf('[2차면접]') >= 0) time = '2차';
      if (ev.getTitle().indexOf('[컬처핏면접]') >= 0) time = '컬처핏면접';
      let creator = ev.getCreators(); // 생성자
      // 캘린더 참석자
      let guest = ev.getGuestList().map(function (d) { return { 'email': d.getEmail(), 'name': getUserNameByEmail(d.getEmail()) } }); // 참여자
      // 이름이 빈경우 제외
      guest = guest.filter(function(d){ return d.name != ''; })
      // var guest = ev.getGuestList().map(function (d) { return { 'email': d.getEmail(), 'name': d.getName() } }); // 참여자
      const desc = ev.getLocation()
      calendarUserList.push({ 'time':time, 'startTime': ev.getStartTime(), 'endTime': ev.getEndTime(), 'title': ev.getTitle(), 'guests': guest, 'desc':desc, 'event': ev })
    }
  }

  /* ---------------------------------------------
     시트의 정보를 가져와 정리한다.
  --------------------------------------------- */
  // 시트의 마지막 row번호를 가져오지만 돌리는건 지정된 라인가지만 돌린다.
  let lr = spreadsheet.getLastRow();
  // 전체를 가져오고 싶다면 아래 한줄을 지우시면 됩니다.
  lr = maxrow;
  // 시트의 데이터를 한번에 퍼온다. (개별로 퍼오면 API호출수 때문에 문제 생김)
  const rows = spreadsheet.getRange('A'+startRow+':X' + lr + '').getValues();
  const sheetUsers = []
  for (let x = 0; x < rows.length; x++) {
    const row = rows[x];
    const status = row[2]; // 지원상태
    const name = row[3]; // 이름
    if (!name) continue; // 이름이 없으면 처리 안함

    const time1 = row[cPos.time1['면접일정']]; // 1차 면접일
    const interviewer1 = row[cPos.time1['면접관']]; // 1차 면접관
    const result1 = row[cPos.time1['결과']]; // 1차 결과
    const notifi1 = row[cPos.time1['불합격 안내']]; // 1차 공유 여부

    const time2 = row[cPos.time2['면접일정']]; // 2차 면접일
    const interviewer2 = row[cPos.time2['면접관']]; // 2차 면접관
    const result2 = row[cPos.time2['결과']]; // 2차 결과
    const notifi2 = row[cPos.time2['불합격 안내']]; // 2차 공유 여부

    const time3 = row[cPos.time3['면접일정']]; // 컬쳐핏 면접일
    const interviewer3 = row[cPos.time3['면접관']]; // 컬쳐핏 면접관
    const result3 = row[cPos.time3['결과']]; // 컬쳐핏 결과
    const notifi3 = row[cPos.time3['불합격 안내']]; // 컬쳐핏 공유 여부
    // 캘린더와 비교할때 편하게 쓸수있게 값을 정리함
    sheetUsers.push({ 
      'name': name, 'status': status, 
      'time1': {'result': result1, 'interviewer': interviewer1, 'notifi': notifi1, 'startTime': time1},
      'time2': {'result': result2, 'interviewer': interviewer2, 'notifi': notifi2, 'startTime': time2},
      'time3': {'result': result3, 'interviewer': interviewer3, 'notifi': notifi3, 'startTime': time3},
    })
  }

  // 시트를 기준으로 캘린더를 찾아서 일치하면 날짜를 갱신함
  for (let i = 0; i < sheetUsers.length; i++) {
    const calendarUsers = calendarUserList.filter(function (d) { return d.title.indexOf(sheetUsers[i].name) >= 0})
    if (calendarUsers) {
      // Logger.log(calendarUsers)
      // 시트에 날짜 등록
      let rowNum = i + startRow
      for (let z = 0; z < calendarUsers.length; z++) {
        const calendarUser = calendarUsers[z]
        /*-------------- 1차 결과 비교 ---------------*/
        if (calendarUser.time == '1차') {
          // 면접일자 변경
          const sheetInfo = sheetUsers[i].time1
          if (dateFormattor(calendarUser.startTime)+'' != sheetInfo.startTime+'') {
            Logger.log('1차 면접일 변경 r='+rowNum+ ' name='+ calendarUser.title+ '/'+calendarUser.startTime+'=>'+sheetInfo.startTime)
            spreadsheet.getRange(columnToLetter(cPos.time1['면접일정']+1) + rowNum).setValue(dateFormattor(calendarUser.startTime))
          }
          // 면접관 갱신 
          const guestsStr = calendarUser.guests.map(function(d) { return d.name }).join(', ')
          if (guestsStr != sheetInfo.interviewer) {
            Logger.log('1차 면접관 변경 r='+rowNum+ ' name='+sheetUsers[i].name + ' '+sheetInfo.interviewer +'=>'+guestsStr)
            spreadsheet.getRange(columnToLetter(cPos.time1['면접관']+1) + rowNum).setValue(guestsStr)
          }
          // 결과와 캘린더 내용이 일치 하지 않는다면 갱신함
          const desc = "결과="+sheetInfo.result+', 불합격 안내='+sheetInfo.notifi
          if (calendarUser.desc.indexOf(desc) == -1) {
            Logger.log('1차 결과 변경  '+sheetInfo.name + calendarUser.title+ '/'+calendarUser.desc+'=>'+desc)
            calendarUser.event.setLocation(desc)
          }
        }
        /*-------------- 2차 결과 비교 ---------------*/
        if (calendarUser.time == '2차') {
          // 면접일자 변경
          const sheetInfo = sheetUsers[i].time2
          if (dateFormattor(calendarUser.startTime)+'' != sheetInfo.startTime+'') {
            Logger.log('2차 면접일 변경 r='+rowNum+ ' name='+ calendarUser.title)
            spreadsheet.getRange(columnToLetter(cPos.time2['면접일정']+1) + rowNum).setValue(dateFormattor(calendarUser.startTime))
          }
          // 면접관 갱신 
          const guestsStr = calendarUser.guests.map(function(d) { return d.name }).join(', ')
          if (guestsStr != sheetInfo.interviewer) {
            Logger.log('2차 면접관 변경 r='+rowNum+ ' name='+sheetUsers[i].name + ' '+sheetInfo.interviewer +'=>'+guestsStr)
            spreadsheet.getRange(columnToLetter(cPos.time2['면접관']+1) + rowNum).setValue(guestsStr)
          }
          // 결과와 캘린더 내용이 일치 하지 않는다면 갱신함
          const desc = "결과="+sheetInfo.result+', 불합격 안내='+sheetInfo.notifi
          if (calendarUser.desc.indexOf(desc) == -1) {
            Logger.log('2차 결과 변경  '+sheetUsers[i].name + '/'+calendarUser.desc+'=>'+desc)
            calendarUser.event.setLocation(desc)
          }
        }
        /*-------------- 컬쳐핏 결과 비교 ---------------*/
        if (calendarUser.time == '컬처핏면접') {
          // 면접일자 변경
          const sheetInfo = sheetUsers[i].time3
          if (dateFormattor(calendarUser.startTime)+'' != sheetInfo.startTime+'') {
            Logger.log('컬쳐핏 면접일 변경 r='+rowNum+ ' name='+ calendarUser.title)
            spreadsheet.getRange(columnToLetter(cPos.time3['면접일정']+1) + rowNum).setValue(dateFormattor(calendarUser.startTime))
          }
          // 면접관 갱신 
          const guestsStr = calendarUser.guests.map(function(d) { return d.name }).join(', ')
          if (guestsStr != sheetInfo.interviewer) {
            Logger.log('컬쳐핏 면접관 변경 r='+rowNum+ ' name='+sheetUsers[i].name + ' '+sheetInfo.interviewer +'=>'+guestsStr)
            spreadsheet.getRange(columnToLetter(cPos.time3['면접관']+1) + rowNum).setValue(guestsStr)
          }
          // 결과와 캘린더 내용이 일치 하지 않는다면 갱신함
          const desc = "결과="+sheetInfo.result+', 불합격 안내='+sheetInfo.notifi
          if (calendarUser.desc.indexOf(desc) == -1) {
            Logger.log('컬쳐핏 결과 변경  '+sheetUsers[i].name + '/'+calendarUser.desc+'=>'+desc)
            calendarUser.event.setLocation(desc)
          }
        }
      }
    } else {
      // 일치 유저 없음
    }
  }
}

/**--------------------------------------------
 * 랩스 동기화용 실행 함수 
 --------------------------------------------*/
function syncLabs() {
  // 랩스 동기화
  const calendarId = 'c_tlm3uu6hs8jidf3h4si6amsfss@group.calendar.google.com'; // 실제 채용 캘린더 ID
  // var calendarId = 'c_bdhuq9fcdsvd3nb9gitihme910@group.calendar.google.com'; // 테스트 채용
  goCandydateSyncCalendar('Candidates_labs', calendarId, 60) // <- 랩스는 60줄까지만 글음
}
/**--------------------------------------------
 * 사업부 동기화용 실행 함수 
 --------------------------------------------*/
function syncBuz() {
  // 사업부 동기화
  const calendarId = 'c_tlm3uu6hs8jidf3h4si6amsfss@group.calendar.google.com'; // 실제 채용 캘린더 ID
  // var calendarId = 'c_bdhuq9fcdsvd3nb9gitihme910@group.calendar.google.com'; // 테스트 채용
  goCandydateSyncCalendar('Candidates', calendarId, 165) // <- 사업부서는 165줄까지만 글음
}
/**--------------------------------------------
 * 스프레드 시트에 매뉴 추가
 --------------------------------------------*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("지금 동기화하기")
    .addItem('랩스 지금 동기화하기', 'syncLabs')
    .addItem('사업부 지금 동기화하기', 'syncBuz')
    .addItem('테스트', 'apiTest')
   // .addItem('캘린더 동기화', 'getSyncCalendarToSummary')
    .addToUi();
}
```
