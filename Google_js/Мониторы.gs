function myif(arg_filter, arg_value){
// проверка условий для фильтра в мниторах
    if (!arg_filter){
      return true
    }
    else if (arg_filter == arg_value){
      return true
    }
    else { return false}
}

function mcrs_update_price_monitor() {
// выполнеятся через меню
// обновляет монитор цен

  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  let max_row = sheet.getMaxRows()
  // считаем диапазон цен
  let adr_map = find_price_range(sheet)
  let all_price = sheet.getSheetValues(20, 40, max_row - 19 , 1 )
  let row = 19

  for (let i=0; i<all_price.length; i++){
    let value = all_price[i][0]
    row++
    set_priceinmonitor(sheet, row, value, adr_map)
  }
  // очищаем монитор
  sheet.getRange(15, 62, 1, 15).clearContent()
  //

  let filter_arr = sheet.getSheetValues(6, 63, 5, 1)

  //let data_arr = sheet.getSheetValues(20, 8, max_row-19, (41-8+1) )
  let data_arr = sheet.getSheetValues(20, 8, max_row-19, (38-8) )
  let data_arr_1 = sheet.getSheetValues(20, 40, max_row-19, 1)

  let column_strat_group = 8
  let column_project = 9
  let column_pitomec = 19
  let column_klient = 33
  let column_brend = 34
  let column_price = 40

  let monitor =  new Array()
  let sum_count_noprice = 0
  let sum_count_price = 0

  for (let i=0; i<adr_map.length; i++){
    monitor.push(new Array())
    monitor[i][0] = 0 // счетчик всего
    monitor[i][2] = adr_map[i][2]
  }

  for (let row=0; row<data_arr.length; row++){
    let strat_group = data_arr[row][column_strat_group-8]
    let project = data_arr[row][column_project-8]
    let pitomec = data_arr[row][column_pitomec-8]
    let klient = data_arr[row][column_klient-8]
    let brend = data_arr[row][column_brend-8]

    let price = data_arr_1[row][0]



    if ( myif(filter_arr[0][0],strat_group) && myif(filter_arr[1][0],project)  &&
         myif(filter_arr[2][0],pitomec)  && myif(filter_arr[3][0],klient)  &&
         myif(filter_arr[4][0],brend) ){
           // найти в какой дипазон попадает цена
          let find = false
          for ( let i=0; i<adr_map.length; i++){
            //  Logger.log("Set:" + value + " s" + adr_map[i][0] + " e" + adr_map[i][1] + " a" + adr_map[i][2] )
            if (price >=adr_map[i][0] && price<=adr_map[i][1] ){
               monitor[i][0]++
               find = true
              }
            }
          if (find){ sum_count_price++}
          else {sum_count_noprice++ }
       }
  }
  let sum_count = sum_count_price + sum_count_noprice

  for ( let i=0; i<adr_map.length; i++){
    count = monitor[i][0]
    sheet.getRange(15, monitor[i][2]).setValue(count)
  }
  sheet.getRange(12, 63).setValue(sum_count)

  if (sum_count>0){
    let procent_price = (sum_count_price)/sum_count
    sheet.getRange(13, 63).setValue(procent_price)
  }
  else {sheet.getRange(13, 63).clearContent() }
}

function mcrs_update_plan_monitor(){
// выполняется из меню
// обновляет монитор планирования

  function row_in_filter(){
  // возвращает массив строк подходящих под фильтр
    let filter_arr = sheet.getSheetValues(7, 95, 6, 1)
    let data_arr = sheet.getSheetValues(20, 8, max_row-19, (28-8+1) )
    let column_for_filter = new Array()
    column_for_filter[0] = 8 //column_strat_group = 8
    column_for_filter[1] = 9 //column_project = 9
    column_for_filter[2] = 10 // column_owner= 10
    column_for_filter[3] = 26 // column_tehnology = 26
    column_for_filter[4] = 27 // column_group = 27
    column_for_filter[5] = 28 // column_line = 28

    let return_arr = new Array()
    let i = 0
    for (let row=0; row<data_arr.length; row++){
      let filter_ok = true
      for (let i =0; i<=5; i++){
          filter_ok = filter_ok && myif(filter_arr[i][0],data_arr[row][column_for_filter[i]-8])
      }
      if ( filter_ok){
           return_arr[i]=row
           i++
      }
    }
    // установим фильтр, если в мониторе стоит фильтр и если нужно установить фильтр
      let yes_filter = sheet.getSheetValues(13, 95,1,1)[0][0]
      if (yes_filter == "Да"){
        let my_filter = sheet.getRange(19, 8, max_row-18, 120-8 ).getFilter()
        if (my_filter){
          my_filter.remove()
        }
        let have_filter_criterii = false
        for (let i=0; i<=5; i++){
          if (filter_arr[i][0]){
            have_filter_criterii = true
          }
        }
        if (have_filter_criterii){
          my_filter = sheet.getRange(19, 8, max_row-18, 120-8 ).createFilter()
          for (i=0; i<=5; i++){
            if (filter_arr[i][0]){
              let builder = SpreadsheetApp.newFilterCriteria()
              builder.whenTextContains(filter_arr[i][0])
              let my_criteria = builder.build()
              my_filter.setColumnFilterCriteria(column_for_filter[i], my_criteria)
            }
          }
        }
      }
    return return_arr
  }
  function findfriday(arg_date){
  // функция ищет ближаюшую пятницу по переданной дате субботу и воскресненье
  // возвоащает назад все остальные дни вперед
    let now_weekday = arg_date.getDay()
    if (now_weekday == 5){
      return arg_date
    }
    let MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    time_date = arg_date.getTime()
    if (now_weekday == 6) { time_date = time_date - MILLIS_PER_DAY}
    else if (now_weekday == 0 ) { time_date = time_date - 2*MILLIS_PER_DAY}
    else {
      let count = 0
      count = 5 - now_weekday
      time_date = time_date + count*MILLIS_PER_DAY
    }
    let new_date = new Date(time_date)
    return new_date
  }

  function month_name(arg_date){
  // массив с русскими названиями месяцев
    let number_month = arg_date.getMonth()
    let name_arr = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
    return name_arr[number_month]
  }

  function draw_table(sheet, column_num, count, previos_column, max_row ){
  // рисует шапку в таблице монитора планирования
    let len_column = column_num+count - previos_column
    if ( len_column !=0) {
      sheet.getRange(17 ,previos_column, 1, len_column )
      .mergeAcross()
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
      sheet.getRange(17 ,previos_column, 3, len_column ).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(20 ,previos_column, max_row-19 , len_column ).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(18 ,previos_column, 1, len_column ).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
      for (let cl = previos_column; cl<=(previos_column+len_column-4); cl = cl + 2 ){
        sheet.getRange(18 ,cl, 2, 2 ).setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
        sheet.getRange(20 ,cl , max_row-19, 2 ).setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
      }
    }
  }

  function getstagecolor() {
  // со страницы этаапы процесса собираем символ этапа и цывет этапа
    let ss = SpreadsheetApp.getActive();
    let main_sheet = "Этапы процесса"
    let sheet = ss.getSheetByName(main_sheet)
    // найдем строку Шаблоны процессов{ и Шаблоны процессов}
    let tag = ["Шаблон{", "Шаблон}"]
    let index = find_tag(tag, sheet)
    let start_row = index.get(tag[0])[0]+2
    let end_row = index.get(tag[1])[0]-1
    let return_arr = new Array()
    let count = 0
    for (let i = start_row; i <= end_row; i++){
      let value = sheet.getRange(i, 10).getValue()
      let bac_color = sheet.getRange(i, 10).getBackground()
      if (value){
        return_arr.push(new Array())
        return_arr[count][0] = value
        return_arr[count][1] = bac_color
        count++
      }
    }
    return return_arr
  }

  function isfinish(arg_row, arg_column, stage_finish_color, stage_finish_font, template_color, template_font ){
  //проверяет по цвету и шрифт что этап выполнен
    let return_value = false
    if (stage_finish_color[arg_row][arg_column] == template_color[0][0]
    && stage_finish_font[arg_row][arg_column] == template_font[0][0] ){
      return_value = true
    }
    return return_value

  }
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  let max_row = sheet.getMaxRows()
  // отчистка пердыдущего офорлмения
    sheet.getRange(17, 90, 1, 30)
      .breakApart()
      .clear({contentsOnly: true, skipFilteredRows: true})
      .setBorder(false, false, false, false, false, false)
    sheet.getRange(18, 90, max_row - 17 , 30).setBorder(false, false, false, false, false, false)
    sheet.getRange(19, 90,1 , 30).setBackground("#cfe2f3")
    sheet.getRange(20, 88, max_row - 19 , 33)
      .setBackground(null)
    . clearContent()

  // расставим даты в шапку монитора
    let MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    let date_arr = []
    let sheet_arr = []
    sheet_arr[0] = new Array()
    let now = new_date()
    let start_date = sheet.getRange(6,95).getValue()
    if (!start_date){
      start_date = now
      sheet.getRange(6,95).setValue(start_date)
    }
    let my_count = 0
    sheet.getRange(19,88, 1, 1).setValue(now)
    let now_time = now.getTime()
    let now_week_kolumn = null
    for (let i = 0; i<=14; i++ ){
      let my_date = findfriday(start_date)
      let my_date_time = my_date.getTime()
      date_arr[i] = my_date
      let sunday = my_date_time + 2 * MILLIS_PER_DAY
      let monday = my_date_time - 4 * MILLIS_PER_DAY
      start_date = new Date(sunday + MILLIS_PER_DAY)
      if ( now_time >= monday && now_time <= sunday ){
        now_week_kolumn = i
      }
      sheet_arr[0][my_count] = my_date
      sheet_arr[0][my_count+1] = ""
      my_count = my_count + 2
    }
    sheet.getRange(19,90, 1, 30 ).setValues(sheet_arr)
    if (now_week_kolumn != null ) {
      sheet.getRange(19,90 + now_week_kolumn*2, 1, 1 ).setBackground("#6fa8dc")
    }
    // расставим месяца и нарисуем талицк
      let previos_month = ""
      let column_num = 88
      let count = 2
      let previos_column = 90
      for (let i = 0; i<=14; i++ ){
        let name_month = month_name(date_arr[i])
        if (name_month != previos_month){
          previos_month = name_month
          sheet.getRange(17 ,column_num+count).setValue(name_month)
          draw_table(sheet, column_num, count, previos_column, max_row )
          previos_column = column_num+count
          column_num = column_num + count
        }
        else { column_num = column_num + count}
      }
    draw_table(sheet, column_num, count, previos_column, max_row )
  // из зоны планирования этапов забираем даты и нарисуем этапы в мониторе
    let my_row = row_in_filter()
    let collor_symbol = getstagecolor()
    let stage_date = sheet.getSheetValues(20, 76, max_row - 19, 8)
    let stage_finish_color = sheet.getRange(20, 76, max_row - 19, 8).getBackgrounds()
    let stage_finish_font = sheet.getRange(20, 76, max_row - 19, 8).getFontColors()
    let template_color = sheet.getRange(2, 74, 1, 8).getBackgrounds()
    let template_font = sheet.getRange(2, 74, 1, 8).getFontColors()

    let standart_incr = sheet.getSheetValues(18, 76, 1, 8)
    let monitor =new Array()
    let monitor_backgrounds = new Array()
    let monitor_font = new Array()
    let count_stge = new Array
    count_stge[0] = new Array()

    for (let i=0; i<=max_row-19-1;i++){
      monitor.push(new Array())
      monitor_backgrounds.push(new Array())
      monitor_font.push(new Array())
      for (let i1=0; i1<=32;i1++){
        monitor[i][i1]=""
        monitor_backgrounds[i][i1]=null
        monitor_font [i][i1] = template_font[0][7]
        count_stge[0][i1]=0
      }
    }

    let return_index = getholiday(ss)
    let epoch_holiday_arr = return_index[0]
    let dayoff_holiday_arr = return_index[1]
    let start_0_week = date_arr[0].getTime() - 4 * MILLIS_PER_DAY// понедельник
    let end_14_week = date_arr[14].getTime() + 2 * MILLIS_PER_DAY// воскресень

    for (let i = 0; i<stage_date.length; i++){ // перебираем строки из массива с этапами
      if (!my_row.includes(i)){ continue}
      for (let cl = 0; cl <= 8; cl ++){ // перебираем колонки от массива с этапами - cfvb 'nfgs
        let end_stage = stage_date[i][cl]
        let incr_in_row = standart_incr[0][cl]
        let my_incriment
        if (end_stage && typeof(end_stage) == "string"){
          let my_index = string_tovalue(end_stage)
          end_stage = my_index[0]
          my_incriment = my_index[1]
          if (!my_incriment) {my_incriment = incr_in_row }
        }
        else { my_incriment = incr_in_row}

        if (!end_stage){continue}
        let start_stage = add_date(end_stage ,-1* (my_incriment-1) ,epoch_holiday_arr, dayoff_holiday_arr )
        end_stage = end_stage.getTime()
        start_stage = start_stage.getTime()
        let column_monitor = 0
        if ( start_stage < start_0_week || end_stage < start_0_week ) {
          monitor[i][1] = "<<<"
          count_stge[0][0]++
        }
        if (start_stage > end_14_week || end_stage > end_14_week ) {
          monitor[i][32] = ">>>"
          count_stge[0][32]++
        }
        if ( now_time >= start_stage && now_time <= end_stage) {
          if (!isfinish(i, cl, stage_finish_color, stage_finish_font, template_color, template_font )){
             monitor[i][0] = collor_symbol[cl][0]
            monitor_backgrounds[i][0] = collor_symbol[cl][1]
          }

        }

        for (let dt =0; dt <=14; dt++){ // для каждого этапа перибераем даты в мониторе
          column_monitor = column_monitor + 2
          let end_week = date_arr[dt].getTime() + 2 * MILLIS_PER_DAY// воскресень
          let start_week = date_arr[dt].getTime() - 4 * MILLIS_PER_DAY// понедельник
          if ( (start_stage <= start_week && end_stage >= end_week) || ( start_stage >= start_week && start_stage <= end_week )
                || ( end_stage >= start_week && end_stage <= end_week) ) {
            count_stge[0][column_monitor]++
            if (monitor[i][column_monitor] == 0) {
              monitor[i][column_monitor] = collor_symbol[cl][0]
              monitor[i][column_monitor+1] = collor_symbol[cl][0]

              if (isfinish(i, cl, stage_finish_color, stage_finish_font, template_color, template_font )) {
                monitor_backgrounds[i][column_monitor] = template_color[0][0]
                monitor_backgrounds[i][column_monitor+1] = template_color[0][0]
                monitor_font[i][column_monitor] = template_font[0][0]
                monitor_font[i][column_monitor+1] = template_font[0][0]
              }
              else {
                monitor_backgrounds[i][column_monitor] = collor_symbol[cl][1]
                monitor_backgrounds[i][column_monitor+1] = collor_symbol[cl][1]
              }
            }
            else if (monitor[i][column_monitor+1] == monitor[i][column_monitor]) {
              monitor[i][column_monitor+1] = collor_symbol[cl][0]
              if (isfinish(i, cl, stage_finish_color, stage_finish_font, template_color, template_font )) {
                monitor_backgrounds[i][column_monitor+1] = template_color[0][0]
                monitor_font[i][column_monitor+1] = template_font[0][0]
              }
              else {
                monitor_backgrounds[i][column_monitor+1] = collor_symbol[cl][1]
                monitor_font[i][column_monitor+1] = template_font[0][7]
              }
            }
          }
        }
      }
    }
    sheet.getRange(20, 88, max_row - 19, 33)
    .setValues(monitor)
    .setBackgrounds(monitor_backgrounds)
    .setFontColors(monitor_font)
// выводим количество этапов
  sheet.getRange(14, 88,1, 33).setValues(count_stge)








}

