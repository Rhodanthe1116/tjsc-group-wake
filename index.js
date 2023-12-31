function splitDataIntoRandomGroups(N = 10, SHEET = '表單回應 1') {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET);
  var allPeople = sheet.getDataRange().getValues().slice(1);


  const NAME = 2
  const NATION = 3
  const GENDER = 4
  const FIRST_EVENT = 9
  const OK = allPeople[0].length
  const OK_S = '可以參加/参加できます'
  const NOTSURE = allPeople[0].length + 1
  const NOTSURE_S = '想參加不確定有沒有空/参加したいのですが現時点日程未定'
  const NO = allPeople[0].length + 2
  const NO_S = '無法參加/参加できない'


  for (var i = 0; i < allPeople.length; i++) {
    allPeople[i].push(allPeople[i].filter(val => val === OK_S).length)
    allPeople[i].push(allPeople[i].filter(val => val === NOTSURE_S).length)
    allPeople[i].push(allPeople[i].filter(val => val === NO_S).length)
  }
  // Shuffle the data randomly
  allPeople = shuffleArray(allPeople);
  // allPeople.sort((p1, p2) => (p1[OK] + p1[NOTSURE] * 0.9) - (p2[OK] + p2[NOTSURE] * 0.9))
  // insertSheet(people, '計算過後')

  const japanese = allPeople.filter(isJp)
  const taiwanese = allPeople.filter(isTw)
  console.log(japanese.length)
  console.log(taiwanese.length)

  const EVENT_NUM = 6
  const MAX_NUM = 72
  const MAX_NUMS = [
    72, 64, 64, 64, 64, 72
  ]
  const peopleOfEvents = []
  for (let i = 0; i < EVENT_NUM; i++) {
    peopleOfEvents.push([])
  }

  // 每場大約max_num人
  // 每個人參加次數盡量平均MAX_NUM * EVENT_NUM / people.length
  // 每場日台各一半
  // 男女平均分散在各個場次
  const expectedTimes = Math.round(MAX_NUM * EVENT_NUM / allPeople.length)
  console.log({ expectedTimes })

  const OKEVENT = allPeople[0].length
  for (var i = 0; i < allPeople.length; i++) {
    const okEvents = []
    for (let event = 0; event < EVENT_NUM; event++) {
      if (allPeople[i][FIRST_EVENT + event] === OK_S) {
        okEvents.push(event)
      }
    }
    allPeople[i].push(okEvents)
  }
  const NOTSUREEVENT = allPeople[0].length
  for (var i = 0; i < allPeople.length; i++) {
    const notsureEvents = []
    for (let event = 0; event < EVENT_NUM; event++) {
      if (allPeople[i][FIRST_EVENT + event] === NOTSURE_S) {
        notsureEvents.push(event)
      }
    }
    allPeople[i].push(notsureEvents)
  }

  const howMany = {}
  for (let i = 0; i < allPeople.length; i++) {
    howMany[allPeople[i][NAME]] = 0
  }

  for (let t = 0; t < MAX_NUM; t++) {

    if (t === 0) {
      allPeople.sort((p1, p2) => (p1[OK]) - (p2[OK]))
    } else {
      allPeople = shuffleArray(allPeople)
    }

    for (var i = 0; i < allPeople.length; i++) {
      const person = allPeople[i]

      if (person[OKEVENT].length === 0 && howMany[person[NAME]] === 0) {
        person[NOTSUREEVENT].sort((a, b) => {
          const aSta = statistics(peopleOfEvents[a])
          const bSta = statistics(peopleOfEvents[b])
          if (isJp(person) && person[GENDER] === '男') {
            return aSta.jpMale - bSta.jpMale
          }
          if (isJp(person) && person[GENDER] === '女') {
            return aSta.jpFemale - bSta.jpFemale
          }
          if (!isJp(person) && person[GENDER] === '男') {
            return aSta.twMale - bSta.twMale
          }
          if (!isJp(person) && person[GENDER] === '女') {
            return aSta.twFemale - bSta.twFemale
          }
        })

        while (person[NOTSUREEVENT].length > 0) {
          const toJoin = person[NOTSUREEVENT].shift()
          const sta = statistics(peopleOfEvents[toJoin])
          if (isJp(person) && sta.jpNum >= MAX_NUMS[toJoin] / 2) {
            continue
          }
          if (!isJp(person) && sta.twNum >= MAX_NUMS[toJoin] / 2) {
            continue
          }

          peopleOfEvents[toJoin].push(person)
          howMany[person[NAME]] += 1
          break

        }
      }

      person[OKEVENT].sort((a, b) => {
        const aSta = statistics(peopleOfEvents[a])
        const bSta = statistics(peopleOfEvents[b])
        if (isJp(person) && person[GENDER] === '男') {
          return aSta.jpMale - bSta.jpMale
        }
        if (isJp(person) && person[GENDER] === '女') {
          return aSta.jpFemale - bSta.jpFemale
        }
        if (!isJp(person) && person[GENDER] === '男') {
          return aSta.twMale - bSta.twMale
        }
        if (!isJp(person) && person[GENDER] === '女') {
          return aSta.twFemale - bSta.twFemale
        }
      })


      for (let jjj = 0; jjj < person[OKEVENT].length; jjj++) {
        const toJoin = person[OKEVENT][jjj]
        const sta = statistics(peopleOfEvents[toJoin])
        if (isJp(person) && sta.jpNum >= MAX_NUMS[toJoin] / 2) {
          continue
        }
        if (!isJp(person) && sta.twNum >= MAX_NUMS[toJoin] / 2) {
          continue
        }

        peopleOfEvents[toJoin].push(person)
        howMany[person[NAME]] += 1
        person[OKEVENT] = [...person[OKEVENT].slice(0, jjj), ...person[OKEVENT].slice(jjj + 1)]
        break
      }


    }


  }


  /////////////////////////////////

  function statistics(people) {
    const jpNum = people.filter(isJp).length
    const twNum = people.length - jpNum
    const maleNum = people.filter(person => person[GENDER] === '男').length
    const femaleNum = people.filter(person => person[GENDER] === '女').length

    const jpMale = people.filter(isJp).filter(p => p[GENDER] === '男').length
    const jpFemale = people.filter(isJp).filter(p => p[GENDER] === '女').length
    const twMale = people.filter(isTw).filter(p => p[GENDER] === '男').length
    const twFemale = people.filter(isTw).filter(p => p[GENDER] === '女').length

    return {
      jpNum, twNum, maleNum, femaleNum, jpMale,
      jpFemale,
      twMale,
      twFemale,
      jptw: jpNum / twNum,
      mf: maleNum / femaleNum
    }
  }




  peopleOfEvents.forEach((people, i) => {
    const sheet = insertSheet(people.map((p, index) => [p[NAME], p[NATION], p[GENDER], p[p.length - 4], p[p.length - 3], p[p.length - 2], p[p.length - 1]]), '場次 ' + (i + 1))
    sheet.getRange(1, 9).setValues([['=query(A:G, "select B, count(A) group by B")']]);
    sheet.getRange(1, 9).setValues([['=query(A:G, "select B, count(A) group by B")']]);
    sheet.getRange(10, 9).setValues([['=query(A:G, "select C, count(A) group by C")']]);
  })

  const howManyTimesFinal = {}
  allPeople.forEach(person => howManyTimesFinal[person[NAME]] = 0)
  peopleOfEvents.forEach((people, i) => {
    people.forEach(person => howManyTimesFinal[person[NAME]] += 1)
  })
  const JOIN_TIMES = allPeople[0].length
  for (var i = 0; i < allPeople.length; i++) {
    allPeople[i].push(howManyTimesFinal[allPeople[i][NAME]])
  }
  console.log('標準差', standardDeviation(allPeople.map(p => p[p.length - 1])))
  console.log('平均', mean(allPeople.map(p => p[p.length - 1])))

  allPeople.sort((p1, p2) => p1[OK] - p2[OK])
  allPeople.sort((p1, p2) => p1[JOIN_TIMES] - p2[JOIN_TIMES])
  const sta = insertSheet(allPeople.map((p, index) => [p[NAME], p[NATION], p[GENDER], p[OK], p[NOTSURE], p[NO],
  `=IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 0}, false) = "無法參加/参加できない", "x", IF(IFNA(VLOOKUP(A${index + 1}, '場次 1'!A:H, 1, false))= A${index + 1}, "O", IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 0}, false) = "想參加不確定有沒有空/参加したいのですが現時点日程未定", "?", "")))`,
  `=IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 1}, false) = "無法參加/参加できない", "x", IF(IFNA(VLOOKUP(A${index + 1}, '場次 2'!A:H, 1, false))= A${index + 1}, "O", IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 1}, false) = "想參加不確定有沒有空/参加したいのですが現時点日程未定", "?", "")))`,
  `=IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 2}, false) = "無法參加/参加できない", "x", IF(IFNA(VLOOKUP(A${index + 1}, '場次 3'!A:H, 1, false))= A${index + 1}, "O", IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 2}, false) = "想參加不確定有沒有空/参加したいのですが現時点日程未定", "?", "")))`,
  `=IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 3}, false) = "無法參加/参加できない", "x", IF(IFNA(VLOOKUP(A${index + 1}, '場次 4'!A:H, 1, false))= A${index + 1}, "O", IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 3}, false) = "想參加不確定有沒有空/参加したいのですが現時点日程未定", "?", "")))`,
  `=IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 4}, false) = "無法參加/参加できない", "x", IF(IFNA(VLOOKUP(A${index + 1}, '場次 5'!A:H, 1, false))= A${index + 1}, "O", IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 4}, false) = "想參加不確定有沒有空/参加したいのですが現時点日程未定", "?", "")))`,
  `=IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 5}, false) = "無法參加/参加できない", "x", IF(IFNA(VLOOKUP(A${index + 1}, '場次 6'!A:H, 1, false))= A${index + 1}, "O", IF(VLOOKUP($A${index + 1}, '表單回應 1'!$C:$P, ${8 + 5}, false) = "想參加不確定有沒有空/参加したいのですが現時点日程未定", "?", "")))`,
  `=COUNTIF(G${index + 1}:L${index + 1}, "O")`
  ]), '參加次數')
  const staArr = [[``, '總人數', '日台比', '男女比', '日本', '台灣', '男生', '女生', '日男', '日女', '台男', '台女']]
  peopleOfEvents.forEach((people, i) => {
    const { jpNum, twNum, maleNum, femaleNum, jpMale, jpFemale, twMale, twFemale } = statistics(people)
    staArr.push([`第${i + 1}場`, jpNum + twNum, jpNum / twNum, maleNum / femaleNum, jpNum, twNum, maleNum, femaleNum, jpMale, jpFemale, twMale, twFemale])
    console.log(`第${i + 1}場`)
    console.log('總人數', people.length)
    console.log('jp / tw', jpNum / twNum)
    console.log('男 / 女', maleNum / femaleNum)
  })
  sta.getRange(1, 15, staArr.length, staArr[0].length).setValues(staArr);


  //////////////////////////////

  peopleOfEvents.forEach((allPeopleOfEvent, eventIndex) => {
    const MAX_GROUP_NUM = 6
    const MAX_GROUP_NUMS = [
      6, 4, 4, 4, 4, 6
    ]
    const GROUP_NUM = allPeopleOfEvent.length / MAX_GROUP_NUMS[eventIndex]
    const peopleOfGroups = []
    for (let i = 0; i < GROUP_NUM; i++) {
      peopleOfGroups.push([])
    }

    const OKGROUP = OKEVENT
    for (let i = 0; i < allPeopleOfEvent.length; i++) {
      const okGroups = []
      for (let group = 0; group < GROUP_NUM; group++) {
        okGroups.push(group)
      }
      allPeopleOfEvent[i][OKGROUP] = okGroups
    }

    for (let i = 0; i < allPeopleOfEvent.length; i++) {

      const person = allPeopleOfEvent[i]
      person[OKGROUP].sort((a, b) => {
        const aSta = statistics(peopleOfGroups[a])
        const bSta = statistics(peopleOfGroups[b])
        if (isJp(person) && person[GENDER] === '男') {
          return aSta.jpMale - bSta.jpMale
        }
        if (isJp(person) && person[GENDER] === '女') {
          return aSta.jpFemale - bSta.jpFemale
        }
        if (!isJp(person) && person[GENDER] === '男') {
          return aSta.twMale - bSta.twMale
        }
        if (!isJp(person) && person[GENDER] === '女') {
          return aSta.twFemale - bSta.twFemale
        }
      })

      while (person[OKGROUP].length > 0) {
        const toJoin = person[OKGROUP].shift()
        const sta = statistics(peopleOfGroups[toJoin])
        if (isJp(person) && sta.jpNum >= MAX_GROUP_NUMS[eventIndex] / 2) {
          continue
        }
        if (!isJp(person) && sta.twNum >= MAX_GROUP_NUMS[eventIndex] / 2) {
          continue
        }

        peopleOfGroups[toJoin].push(person)
        break

      }
    }

    const sheetOfThisEvent = insertSheet([['']], `場次 ${eventIndex + 1} 分組`)
    const JPLEVEL = 6
    const TWLEVEL = 7
    peopleOfGroups.forEach((people, i) => {
      const sta = statistics(people)

      sheetOfThisEvent.getRange(sheetOfThisEvent.getLastRow() + 1, 1, 1, 7).setValues([[`場次 ${eventIndex + 1} - 組別 ${i + 1}`, '日本', '台灣', '日男', '日女', '台男', '台女']]);
      sheetOfThisEvent.getRange(sheetOfThisEvent.getLastRow() + 1, 1, 1, 7).setValues([[``, sta.jpNum, sta.twNum, sta.jpMale, sta.jpFemale, sta.twMale, sta.twFemale]]);

      const arr = people.map((p, index) => [p[NAME], p[NATION], p[GENDER], p[isJp(p) ? TWLEVEL : JPLEVEL]])
      sheetOfThisEvent.getRange(sheetOfThisEvent.getLastRow() + 1, 1, arr.length, arr[0].length).setValues(arr)

      sheetOfThisEvent.getRange(sheetOfThisEvent.getLastRow() + 1, 1, 1, 1).setValues([['  ']]);

    })

  })


  function isJp(person) {
    return person[NATION] === '日本' || person[NATION] === '台日混血/日台ハーフ'
  }
  function isTw(person) {
    return !isJp(person)
  }

  return
}


function mean(arr) {
  const n = arr.length;
  const mean = arr.reduce((a, b) => a + b) / n;
  return mean
}
function standardDeviation(arr) {
  const n = arr.length;
  const _mean = mean(arr);
  const variance = arr.reduce((a, b) => a + (b - _mean) ** 2, 0) / n;
  const stdDev = Math.sqrt(variance);
  return stdDev;
}

function getMultipleRandom(arr, num) {
  const shuffled = [...arr].sort(() => 0.5 - Math.random());

  return shuffled.slice(0, num);
}

function chosen(pro) {
  return Math.random() < pro
}

function insertSheet(arr, name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var itt = spreadsheet.getSheetByName(name);
  if (itt) {
    spreadsheet.deleteSheet(itt);
  }
  const newSheet = spreadsheet.insertSheet(name);
  newSheet.getRange(1, 1, arr.length, arr[0].length).setValues(arr);
  return newSheet
}

// Helper function to shuffle an array
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}
