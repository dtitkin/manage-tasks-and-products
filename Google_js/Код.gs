function onOpen(e){
// системный тригер при открытии страаницы
 var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Обновление мониторов')
      .addItem('Монитор цены', 'mcrs_update_price_monitor')
      .addSeparator()
      .addItem('Монитор этапов', 'mcrs_update_plan_monitor')
      //addItem('Монитор этапов', '')
      //.addSubMenu(ui.createMenu('Sub-menu')
      //   .addItem('Second item', 'menuItem2'))
      .addToUi();

  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  sheet.getRange(6, 95).setValue(new_date())

  // проверяем наличие формул
  verify_formulas()
}


function onChange(e){
// заложил на бдущее тригер для удаления/добавления строк
  Logger.log(e.changeType)


}


function onEdit(e){
// системный тригер срабатывает после изменения значения в ячейке
  let oldvalue = e.oldValue
  let value = e.value
  let currentrange = e.range
  let currentrow = currentrange.getRow()
  let currentcolumn = currentrange.getColumn()
  let app = e.source
  let name_ss = app.getActiveSheet().getName()
  let user_do = e.user.getEmail()
 //Logger.log("Таблица: "+name_ss+"Строка: "+currentrow+" Колонка: "+currentcolumn+" Старое значение: {"+oldvalue+"}"+" Новое значение: {"+value+"}"+" "+typeof(value))


  if (name_ss == "Проекты"){
    if (currentrow >=29 && (currentcolumn == 1 || currentcolumn == 2)){
      project_newrow_date_user(currentrow)
    }
    // коррекция названия проекта
     if (currentrow >=29 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 9)
    }

  }
  else if (name_ss == "Рабочая таблица №1"){
    if (currentrow >=20 && ( currentcolumn == 9 || currentcolumn == 10)){
      worktable_setprojectowner(currentrow)
    }
    else if (currentrow >=20 &&  currentcolumn == 40  && value != oldvalue){
      worktable_shelf_price(currentrow, value)
    }
    else if (currentrow >=20 &&  ( currentcolumn == 54 ||  currentcolumn == 55 || currentcolumn == 56 )){
        my_f = app.getActiveSheet().getRange(currentrow,57).getFormula()
        if (!my_f){
          app.getActiveSheet().getRange(currentrow,57).setFormula("=SUM(BB"+currentrow+":BD"+currentrow+")")
        }

    }
    else if (currentrow >=20 &&  currentcolumn >= 54 && currentcolumn <=56  && value){
      set_priority_formula(currentrow)
    }
    else if (currentrow >=20 &&  currentcolumn >= 45 && currentcolumn <=46  && value){
      set_prognoz_formula(currentrow)
    }
    else if (currentrow == 18 &&  currentcolumn == 38  && value != oldvalue){
      let ui = SpreadsheetApp.getUi()
      ui.alert("Изменение размеров ячеек займет 2-5 секунды");
      set_size(value)
    }
     else if (currentrow == 5 &&  currentcolumn == 38){

      let allert_txt =""

      if (value == "TRUE"){
        allert_txt = 'Скрыть колонки ?\n Это займет несколько секунд. '
        }
      else {
        allert_txt = 'Показать колонки ?\n Это займет несколько секунд. '
        }
      let ui = SpreadsheetApp.getUi()
      let response = ui.alert(allert_txt, ui.ButtonSet.YES_NO);
      if (response == ui.Button.YES) {
       // Logger.log("ДА")
        hide_unhide_column(value)
      }
    }




  // Планирование этапов
    // 1. Проверка на то что данные из Wrike
    arr_column = [76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86 ]
    // Logger.log("Таблица: "+name_ss+"Строка: "+currentrow+" Колонка: "+currentcolumn+" Старое значение: {"+oldvalue+"}"+" Новое значение: {"+value+"}")
    if ( currentrow >=20 && arr_column.includes(currentcolumn)  ){
      let data_from_wrike = app.getActiveSheet().getSheetValues(currentrow, 74 ,1,1 )[0][0]
      if (data_from_wrike) { data_from_wrike = data_from_wrike.toLowerCase()}
      if (data_from_wrike == "w" || data_from_wrike == "n" || data_from_wrike == "p" ){
        SpreadsheetApp.getUi().alert('Актаульные даты во Wrike.\n Для редактирования установите G в GW');
        let num_gooogle_date = 0
        num_gooogle_date = oldvalue
        if (typeof(oldvalue) == "string") {
          if (oldvalue.slice(-2) == ".0"){
            num_gooogle_date = Number(oldvalue)
          }
        }
        app.getActiveSheet().getRange(currentrow, currentcolumn ).setValue(num_gooogle_date)
      }
    }
    else if (currentrow >=20 && currentcolumn == 74 ){
        set_color_font(app.getActiveSheet(), currentrow, value.toLowerCase())
    }
    // автоматические расчеты даты в блоке планирования
    else if (currentrow >=20 && currentcolumn == 75 ){
      let data_from_wrike = app.getActiveSheet().getSheetValues(currentrow, 74 ,1,1 )[0][0].toLowerCase()
      if (data_from_wrike == "w" && value){
        SpreadsheetApp.getUi().alert('Актаульные даты во Wrike.\n Для редактирования установите G в GW');
        app.getActiveSheet().getRange(currentrow, currentcolumn ).setValue(oldvalue)
      }
      else if ((data_from_wrike == "p" || data_from_wrike == "n") && value){
        SpreadsheetApp.getUi().alert('Проект остановлен или на паузе.\n Для редактирования установите G в GW');
        app.getActiveSheet().getRange(currentrow, currentcolumn ).setValue(oldvalue)
      }
      else if (value){
        worktable_auto_calc(currentrow, value)
      }
    }
  }
  else if (name_ss == "Ресурсы"){

   // Питомцы
    if (currentrow >=94 && currentrow<=109 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 19)
    }
    // Упаковки
     if (currentrow >=114 && currentrow<=148 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 42)
    }
    // Технологии
     if (currentrow >=153 && currentrow<=185 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 26)
    }
    // коллекции
     if (currentrow >=190 && currentrow<=222 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 27)
    }
    // линейка
     if (currentrow >=227 && currentrow<=259 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 28)
    }
    // клиент
     if (currentrow >=264 && currentrow<=314 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 33)
    }
    // бренд
     if (currentrow >=319 && currentrow<=351 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 34)
    }
    // Стратегическая группа
     if (currentrow >=21 && currentrow<=45 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 8)
    }
    // Пользователи и группы Wrike
     if (currentrow >=50 && currentrow<=89 &&  currentcolumn == 2 && oldvalue != ""){
      change_resursname(oldvalue, value, 10)
      change_resursname(oldvalue, value, 3, "Проекты", 30)
      change_resursname(oldvalue, value, 8, "Этапы процесса", 20)
      change_resursname(oldvalue, value, 6, "Задачи этапов", 20)
    }
    else if (currentrow >=50 && currentrow<=89 &&  currentcolumn == 3 && oldvalue != ""){
      change_resursname(oldvalue, value, 8, "Этапы процесса", 20)
      change_resursname(oldvalue, value, 6, "Задачи этапов", 20)
    }

  }

  else if (name_ss == "Этапы процесса"){
    //if (currentcolumn == 3 && (value.toLowerCase() == "э" || value == "з" )){
    //  settings_copy(currentrow, value)
    //}
    let arr_column = [1, 3, 7, 10]
    if ( arr_column.includes(currentcolumn)){
      make_head(currentrow)

    }

  }


}

function set_size(argsize){
// меняет размер ячейки с фотографией
  let size = Number(argsize)
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  let min_size_column = 50
  let min_size_row = 21

  let max_size_column = 200
  let max_size_row = 200
  let set_size = 0

  set_size = size
  if (!size || size < min_size_column){set_size = min_size_column}
  else if (size>max_size_column){set_size=min_size_column}
  sheet.setColumnWidth(38, set_size)

  if (!size || size < min_size_column){set_size=min_size_row}
  else if (size > max_size_column){set_size=max_size_row}
  max_size = sheet.getMaxRows()
  for (i=19; i<= max_size; i++){
    sheet.setRowHeight(i, set_size)
  }
}

function hide_unhide_column(arg_bool){
// скрывает или отображаает столбцы
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  let max_column = sheet.getMaxColumns()
  let my_column = sheet.getRange(4, 1, 1, max_column).getValues()
  //Logger.log("Количество колонок " +max_column+ " размер массива"+ my_column[0].length+" аргумент"+arg_bool )
  for (let i=0;i<max_column;i++){
    if (my_column[0][i] == "1"){
      if (arg_bool == "TRUE"){
        let column=sheet.getRange(4, i+1)
        sheet.hideColumn(column)
      }
      else{
        let column=sheet.getRange(4, i+1)
        sheet.unhideColumn(column)
      }

    }

  }


}




function verify_formulas(){
// вызывается при открытии. проверят формулы и устанавливает их
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)

  let max_row = sheet.getMaxRows()
  let formulas_prioity = sheet.getRange(20, 57, max_row - 19, 1).getFormulasR1C1()
  let base_formula = sheet.getRange(2, 57).getFormulaR1C1()
  for (let i = 0; i<formulas_prioity.length; i++){
    if (!formulas_prioity[i][0]){
      sheet.getRange(i+20 , 57).setFormulaR1C1(base_formula)
    }
  }

  formulas_prioity = sheet.getRange(20, 47, max_row - 19, 2).getFormulasR1C1()
  let base_formula_0 = sheet.getRange(2, 47).getFormulaR1C1()
  let base_formula_1 = sheet.getRange(2, 48).getFormulaR1C1()
  for (let i = 0; i<formulas_prioity.length; i++){
    if (!formulas_prioity[i][0]){
      sheet.getRange(i+20 , 47).setFormulaR1C1(base_formula_0)
    }
    if (!formulas_prioity[i][1]){
      sheet.getRange(i+20 , 48).setFormulaR1C1(base_formula_1)
    }

  }






}

function set_priority_formula(row){
// устанавливает формулу в троку если ее там нет
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  let base_formula = sheet.getRange(2, 57).getFormulaR1C1()
  let my_formula = sheet.getRange(row, 57).getFormulaR1C1()
  if (!my_formula) {
    sheet.getRange(row, 57).setFormulaR1C1(base_formula)
  }
}

function set_prognoz_formula(row){
// устанавливает формулу в троку если ее там нет
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  let base_formula_0 = sheet.getRange(2, 47).getFormulaR1C1()
  let base_formula_1 = sheet.getRange(2, 48).getFormulaR1C1()

  let my_formula_0 = sheet.getRange(row, 47).getFormulaR1C1()
  let my_formula_1 = sheet.getRange(row, 48).getFormulaR1C1()

  if (!my_formula_0) {
    sheet.getRange(row, 47).setFormulaR1C1(base_formula_0)
  }
  if (!my_formula_1) {
    sheet.getRange(row, 48).setFormulaR1C1(base_formula_1)
  }
}


function string_tovalue(cell_value ){
// переоводит 20.01.2021(15) в два отдельных занчения дату и номер
  let pos1 = cell_value.indexOf("[")
  let pos2 = cell_value.indexOf("]")
  if (pos1 == -1 || pos2 == -1  ){
    return [null,null]
  }

  let my_date = cell_value.slice(0,pos1)
  let firstpoint = my_date.indexOf(".", 0)
  let secondpoint = my_date.indexOf(".", firstpoint+1)
  let day = my_date.slice(0,firstpoint )
  let month = my_date.slice(firstpoint+1, secondpoint )
  let year =  my_date.slice(secondpoint + 1 )
  if (year.length==2){year="20"+year}

  let my_number = cell_value.slice(pos1 +1 ,pos2)
  if (firstpoint == -1 || secondpoint == -1  ){
    return [null,my_number]
  }
  day = Number(day)
  month = Number(month)-1
  year = Number(year)
  my_date = new Date(year, month,day )
  my_number = Number(my_number)


  return [my_date, my_number ]

}
function value_tostring(arg_date, arg_num ){
// переводит два значения в строку 20.01.2021(15)
  let day = arg_date.getDate()
  let str_day = String(day)
  if (str_day.length == 1) { str_day = "0"+str_day}
  let month = arg_date.getMonth()+1
  let str_month = String(month)
  if (str_month.length == 1) { str_month = "0"+str_month}

  let year =  arg_date.getFullYear()

  let my_str = "" + str_day+"." + str_month + "."+year + "[" + arg_num + "]"

  return my_str
}

function new_date(){
// создает новую дату но со временем 00:00:00
  let new_d = new Date()
  let new_d00 = new Date(new_d.getFullYear(), new_d.getMonth(), new_d.getDate())
  return new_d00
}

function getholiday(ss){
// считаем праздники, переносы
  holiday_sheet = ss.getSheetByName("Рабочий календарь")
  let tag = ["Шаблон{", "Шаблон}"]
  let index = find_tag(tag, holiday_sheet)
  let start_row = index.get(tag[0])[0]
  let end_row = index.get(tag[1])[0]
  start_row  = start_row + 1
  end_row = end_row -1
  let holiday_arr = holiday_sheet.getSheetValues(start_row, 1,end_row - start_row + 1,2 )
  // сохраним праздники и сохраним перенос дат в массивы со значениями getTime()
  let epoch_holiday_arr = []
  let dayoff_holiday_arr = []
  for (let i=0; i < holiday_arr.length; i++){
    epoch_holiday_arr.push(holiday_arr[i][0].getTime())
    if (holiday_arr[i][1]){
      dayoff_holiday_arr.push(holiday_arr[i][1].getTime())
    }

  }

  return [epoch_holiday_arr, dayoff_holiday_arr]
}

function set_color_font(sheet, row, calc, column_start =0 , column_end =0  ){
// раскрашивает ячейки в планировании этапов в соответствии с шаблоном заданным в верхних скрытых строках
  //считаем шаблон
  // тест
   // let sheet = SpreadsheetApp.getActiveSheet()
   // let row = 20
   // let calc ="g"
   // let column_start =0
   // let column_end =0
  // тест

  let font_template = sheet.getRange(2, 74, 1, 7).getFontColors()
  let color_template = sheet.getRange(2, 74, 1, 7).getBackgrounds()
  if ( calc == "f" ){
    sheet.getRange(row, 76 + column_start -1 , 1,column_end - column_start +1  )
    .setBackground(color_template[0][0])
    .setFontColor(font_template[0][0])
    return
  }
  if ( calc == "x" ){
    sheet.getRange(row, 76 + column_start -1 , 1,column_end - column_start +1  )
    .setBackground(color_template[0][1])
    .setFontColor(font_template[0][1])
    return
  }

  let set_color = new Array()
  set_color[0] = new Array()
  let  set_font = new Array()
  set_font[0] = new Array()
  let now = new_date().getTime()

  let stage_date = sheet.getSheetValues(row, 76, 1, 9)
  let stage_color = sheet.getRange(row, 76, 1, 9).getBackgrounds()
  for (let cl = 0; cl <= 8; cl ++){ // перебираем колонки от массива с этапами
    set_color[0][cl] = color_template[0][1]
    set_font[0][cl] = font_template[0][1]

    let end_stage = stage_date[0][cl]
    let personal_len = false
    let prophesied = false
    if (end_stage && typeof(end_stage) == "string"){
      let my_index = string_tovalue(end_stage)
      end_stage = my_index[0]
      personal_len = true
    }
    end_stage = end_stage.getTime()
    if (end_stage < now){prophesied = true }
    // проверяем что стоял выполненный цвет
    let completed = false
    if (stage_color[0][cl] == color_template[0][0] ){
      set_color[0][cl] = color_template[0][0]
      set_font[0][cl] = font_template[0][0]
      completed= true
    }
    else {
      if (prophesied) { set_font[0][cl] = font_template[0][2]}
      if ( personal_len && !prophesied){ set_font[0][cl] = font_template[0][5]}
      if (calc == "w" && !completed ) {set_color[0][cl] = color_template[0][3]}
      else if (calc == "g" && !completed ) {set_color[0][cl] = color_template[0][1]}
      else if (( calc == "p" || calc == "n") && !completed ) {
        set_color[0][cl] = color_template[0][6]
        set_font[0][cl] = font_template[0][6]
      }
      else if (calc == "calc" && !completed ) {
        if (cl >= column_start && cl <= column_end){
          set_color[0][cl] = color_template[0][4]
        }
        else {
          set_color[0][cl] = color_template[0][1]
        }
      }
    }
  }
  sheet.getRange(row, 76, 1, 9)
  .setBackgrounds(set_color)
  .setFontColors(set_font)
}

function worktable_auto_calc( row, value) {
// выполняет автоматический пересчет дат в разделе планирование этапов
// заполняет строку этапов датами или строками с указанием сохраненных дат
  //let row = 20
  //let value ="D"

  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)

  // определим направление расчета и номер от которого считаем
  let direction_string = value.slice(0, 1)
  direction_string = direction_string.toLowerCase()
  if (direction_string == "d" || direction_string == "r" || direction_string == "f" || direction_string == "x" ){}
  else{ return }

  if (direction_string == "f" || direction_string == "x") {
  // закрашиваем выполненным или убираем этот признак и выходим
    if (!(value.length >= 2) ) {return}
    let pos1 = value.indexOf("-")
    let s_c = 0
    let e_c = 0
    if (pos1 == -1) {
      s_c = Number(value.slice(1))
      e_c = s_c
      }
    else {
      s_c = Number(value.slice(1,pos1))
      e_c = Number(value.slice(pos1+1)) }
      if (s_c >=1 && s_c <=8 && e_c >= 1 && e_c <= 8 ) {
        set_color_font(sheet, row, direction_string, s_c, e_c )
      }
   return
  }

  let direction = 0
  let start_number = 0
  let end_number = 0
  let number_slice = 1
  let save = "S" // режим сохранения значений установлен по умолчанию
  start_number = Number(value.slice(number_slice))-1

  if (direction_string == "d") { end_number = 8 }
  else { end_number =0 }
  if (start_number == -1 ){
    if (direction_string == "d") { start_number = 0 }
    else { start_number =8 }
  }
  if (direction_string == "d") { direction = 1 }
  else {direction = -1}
  // найдем  выполенные ячейки и изменим  start_number и end_number
    let font_template = sheet.getRange(2, 74, 1, 1).getFontColors()
    let color_template = sheet.getRange(2, 74, 1, 1).getBackgrounds()
    let fact_font = sheet.getRange( row, 76 , 1, 9).getFontColors()
    let fact_color = sheet.getRange( row, 76 , 1, 9).getBackgrounds()
    for (let i=0; i<=8; i++){
      if ( fact_font[0][i] == font_template[0][0] && fact_color[0][i] == color_template[0][0]) {
        if (direction == 1){
          start_number = i + 1
        }
        else {

          end_number = i + 1
        }
      }
    }
    if (start_number > 8 || end_number > 8 || ( direction == -1 && end_number > start_number)  ){ return }


  // заберем даты из строки таблицы
  // инкрименты и даты соберем в один массив [[дата][длительность]]
  let my_date = []
  let date_onrow_arr = sheet.getSheetValues( row, 76 , 1, 9)
  let incriment_arr = sheet.getSheetValues( 18, 76, 1, 9)

  for (let i =0; i <= 8; i++ ){

    let cell_value = date_onrow_arr[0][i]
    let incriment_value = incriment_arr[0][i]

    my_date.push(new Array())
    my_date[i][0] = null
    my_date[i][1] = null
    my_date[i][2] = null
    if (cell_value && typeof(cell_value) == "string"){
      // если это строка то вытащим из нее дату и инкримент
        let my_index = string_tovalue(cell_value)
        if (my_index[0]) { my_date[i][0] = my_index[0] }
        else {my_date[i][0] = new_date() }
        if (save == "S"){
          my_date[i][1] = false
          my_date[i][2] = my_index[1]
          if ( my_date[i][2] == null ) {
            my_date[i][1] = incriment_value
            my_date[i][2] = false
          }
        }
        else {
          my_date[i][1] = incriment_value
          my_date[i][2] = false
        }
      }
    else {
        if (cell_value && typeof(cell_value) == "object") {my_date[i][0] = cell_value}
        else {my_date[i][0] = new_date() }
        my_date[i][1] = incriment_value
        my_date[i][2] = false
    }

  }



  let return_index = getholiday(ss)
  let epoch_holiday_arr = return_index[0]
  let dayoff_holiday_arr = return_index[1]


  let return_value = []
  let return_value_1 = []
  return_value[0] = new Array
  for (let i =0; i<=8; i++){
    return_value[0][i] = null
    return_value_1[i] = null
  }
  // выполняем расчет
    let repeat = true
   //for (let i = start_number; if (direction == 1) {i <= end_number } else { i >= end_number}; i = i + direction)

    let i = 0
    let end =0
    if ( direction == -1) {
      i = 8
      end =0
    }
    else {
      i = 0
      end = 8
    }


    while (repeat) {
      repeat = false
      let incriment = 0
      let incriment_toString = 0
      let from_row_orhead = 1
      let  personal_len = (save == "S" && my_date[i][2])

      if (personal_len) { from_row_orhead = 2 }
      incriment_toString = my_date[i][from_row_orhead]
      if (direction == 1) { incriment = my_date[i][from_row_orhead] }
      else {
        if (i == 8 ) { incriment = my_date[i][from_row_orhead] }
        else {
          if (my_date[i+1][1]) {incriment = my_date[i+1][1] }
          else { incriment = my_date[i+1][2] }
        }
      }


      if ( (i <= start_number && direction == 1) || (i >= start_number && direction == -1)
        || ( direction == 1 && i >end_number )  || ( direction == -1  && i <end_number )){
        return_value_1[i] = my_date[i][0]
        if (personal_len) { return_value[0][i] = value_tostring(my_date[i][0],incriment_toString)}
        else { return_value[0][i] = my_date[i][0] }
      }

      else {
        let before_date = return_value_1[i-1*direction]
        let now_date = add_date(before_date ,incriment*direction ,epoch_holiday_arr, dayoff_holiday_arr )
        return_value_1[i] = now_date
        if (personal_len) { return_value[0][i] = value_tostring(now_date,incriment_toString )}
        else { return_value[0][i] = now_date }
      }
      if (direction == 1) {
        i++
        if (i <= end) { repeat = true}
      }
      else if (direction == -1) {
        i--
        if (i >= end) { repeat = true}
      }

   }
  // устанавливаем значение в таблицу
  sheet.getRange(row, 76, 1, 9).setValues(return_value)
  // те колонки которые не учавствовали в расчете
  //sheet.getRange(row, 76, 1, 11).setBackground(null)
  let column_start = 0
  let column_end= 0

  if (direction == 1) { column_end = start_number }
  else {
    column_start = 8
    column_end = start_number
  }
   set_color_font(sheet, row,"calc",column_start, column_end )
}


function add_date(sheet_date ,work_day,epoch_holiday_arr, dayoff_holiday_arr ){
// прибавляет или отнимает день с учетом праздников и выходных
  let MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  //let work_day = 1
  let plus_minus
  let add_day
  if (work_day <= 0 ){
    plus_minus = -1
    add_day = - work_day
  }
  else {
    plus_minus = 1
    add_day = work_day
  }

  let sheet_date_time = sheet_date.getTime()
  for (let i=1; i<=add_day; i++){
    let doit  = true
    let weekday
    let my_day
    while (doit){
      sheet_date_time = sheet_date_time + (plus_minus*MILLIS_PER_DAY)
      my_day = new Date(sheet_date_time)
      doit = false
      weekday = my_day.getDay()
      if (( weekday == 0 || weekday == 6) && !dayoff_holiday_arr.includes(sheet_date_time) ){ doit = true }
      if (epoch_holiday_arr.includes(sheet_date_time)){ doit = true }

    }
  }

   let new_date = new Date(sheet_date_time)
   return new_date

}


function make_head(row){
// реагирует на изменения параметров у эатпа на странице Этапы процесса
// и рисует шапку на странице Рабочая таблица № 1
  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Этапы процесса"
  let sheet = ss.getSheetByName(main_sheet)
  // найдем строку Шаблоны процессов{ и Шаблоны процессов}
  let tag = ["Шаблон{", "Шаблон}"]
  let index = find_tag(tag, sheet)
  let start_row = index.get(tag[0])[0]
  let end_row = index.get(tag[1])[0]
  row_stage  = start_row + 2
  if (row >= row_stage && row < end_row ){

    let tag_arr = sheet.getSheetValues( row, 1, 1, 10)
    let color_b = sheet.getRange(row,10 ).getBackground()
    let stage_num = tag_arr[0][0]
    let stage_title = tag_arr[0][2]
    let len_stage =  tag_arr[0][6]
    //Logger.log(stage_num + " "+stage_title + " 1 "+!stage_title + " "+!stage_num+ " "+typeof(stage_num))

    if (stage_title && stage_num && typeof(stage_num) == "number"){
      if (stage_num>=1 && stage_num<=11 ){
        let w_sheet = ss.getSheetByName("Рабочая таблица №1")
        my_range = w_sheet.getRange(18, 76+stage_num-1, 2)
        let value = [[len_stage],[stage_num+ "\n"+stage_title]]
        my_range.setValues(value)
        my_range.setBackground(color_b)
      }
    }
  }
}

function find_tag(tag, sheet, column = 1 ){
// для страниц с настройками ищет тэги и вовзарщает два адреса
// строку первого тега и строку второго
  let max_row = sheet.getMaxRows()
  let tag_arr = sheet.getSheetValues(1, column, max_row, 1)
  let return_index = new Map()

  for (let key in tag_arr) {
    for (let i =0; i<tag.length; i++){
      if (tag_arr[key][0] == tag[i] ){
          if (!return_index.has(tag[i])){
            return_index.set(tag[i], new Array())
          }
          return_index.get(tag[i]).push(Number(key)+1)
      }
    }
  }
  return return_index
}




function change_resursname(oldvalue, value, num_column, main_sheet = "Рабочая таблица №1", start_row = 20){
// срабатывает при коррекции любого ресурса на странице Ресурсы или других страницах
// и меняет это значение во всех строках таблицы переданной в main_sheet
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(main_sheet)
  let max_row = sheet.getMaxRows()
  let tag_arr = sheet.getSheetValues(start_row, num_column, max_row, 1)
  for ( let i=0; i<tag_arr.length; i++){
    //Logger.log(tag_arr[i][0]+" -"+oldvalue)
    if (tag_arr[i][0] == oldvalue) {
      sheet.getRange(start_row+i, num_column).setValue(value)
    }
  }
}

function worktable_setprojectowner(row){
// на странице Рабочая таблица устаналвивает руководителя проекта при выборе проекта
  let ss = SpreadsheetApp.getActiveSheet();
  let column_name_project = 9
  let column_name_owner = 10
  let my_value = ss.getRange(row, column_name_project, 1, 2).getValues()
  let my_project = my_value[0][0]
  let my_owner = my_value[0][1]

  if (my_project != ""){
    let ss_project =  SpreadsheetApp.getActive().getSheetByName("Проекты");
    column_name_project = 2
    let column_owner = 3
    let tag_arr = ss_project.getSheetValues(30 , column_name_project, ss_project.getMaxRows() - 30 + 1, 2)
    let i = 0
    let find_project = false
    let new_owner
    while (i<tag_arr.length && !find_project){
      if (tag_arr[i][0] == my_project){
        find_project = true
        new_owner = tag_arr[i][1]
      }
      i++
    }
    if (find_project){
      if ( my_owner != new_owner){
        ss.getRange(row,column_name_owner ).setValue(new_owner)
      }
    }
  }





}


function find_price_range(sheet){
// возвращает массив с адресами и значениями колонок с ценами
    let tag_arr = sheet.getSheetValues(19, 62, 1, 12 )
    let adr_map = new Array()
    let start_range = 0
    let end_range = 0
    let start_column = 62
    let i1 = 0
    for ( let i=0; i<tag_arr[0].length; i++){
      if (typeof(tag_arr[0][i]) == 'number'){
        end_range = tag_arr[0][i]
        adr_map.push(new Array())
        adr_map[i1][0] = start_range + 0.1
        adr_map[i1][1] = end_range
        adr_map[i1][2] = start_column
        start_column++
        i1++
        start_range = end_range
     }
    }
    return adr_map
  }


function set_priceinmonitor(sheet, row, value, adr_map){
// находит в какую колнку ставить и ставит
    sheet.getRange(row, 62, 1, 12).clearContent()
    for ( let i=0; i<adr_map.length; i++){
      //  Logger.log("Set:" + value + " s" + adr_map[i][0] + " e" + adr_map[i][1] + " a" + adr_map[i][2] )
      if (value >=adr_map[i][0] && value<=adr_map[i][1] ){
      sheet.getRange(row, adr_map[i][2]).setValue(value)
      }
    }

  }

function worktable_shelf_price(row, value){
// приредаткировании цены устанавливает цену в нужную колонку монитора цен

  let ss = SpreadsheetApp.getActive();
  let main_sheet = "Рабочая таблица №1"
  let sheet = ss.getSheetByName(main_sheet)
  // считаем диапазон цен
  let adr_map = find_price_range(sheet)
  set_priceinmonitor(sheet, row, value, adr_map)
}

function project_newrow_date_user(row) {
// при добавлении проекта на странце Проекты устанавливаает дату ввода записи с проектом
  let ss = SpreadsheetApp.getActiveSheet();
  let now = new Date();

  let column_name_project = 2
  let column_date = 6

  let my_value = ss.getRange(row, 1, 1, 7).getValues()
  if (my_value[0][column_name_project-1] !="" && (my_value[0][column_date-1] == "")){
    let values = [[now]];
    ss.getRange(row, column_date, 1, 1).setValues(values)
  }

  //Logger.log("00" +  my_value[0][0] + "01" + my_value[0][1])

}
