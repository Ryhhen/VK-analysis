import sqlite3
import datetime
import re
import xlsxwriter

# Открываем базу данных "ВКонтакте"
with sqlite3.connect('vkim.sqlite') as db:
 cursor = db.cursor()
 query = """ SELECT messages.local_id, messages.vk_id, messages.cnv_msg_id, groups.title, groups.id, messages.time, messages.from_member_type, messages.from_member_id, messages.is_incoming, messages.body, messages.attach, messages.nested FROM messages, groups WHERE (messages.dialog_id * (-1)) = groups.id """

# Результат запроса присваиваем переменной "cursor"
 cursor.execute(query)
# Формируем из результата запроса кортеж кортежей 
 tuple_from_query = cursor.fetchall()

# Завершаем работу с БД (применяем изменения; в данном случае строка кода необязательна) 
 db.commit()

# Открываем Excel-файл для записи результата и выделяем заголовки жирным шрифтом
excel_file = xlsxwriter.Workbook('best_selection_for_public.xlsx')
excel_worksheet = excel_file.add_worksheet()
font_for_headers = excel_file.add_format({'bold':True})
list_of_headers = ['local_id','vk_id','cnv_msg_id','Название группы/канала','ID собеседника','Дата и время (UTC)','Тип отправителя (1 - пользователь, 3 - группа/канал)','ID отправителя','Направление (1 - входящее, 0 - исходящее)','Текст сообщения','Интернет-ссылка на медиа-сообщение','Интернет-ссылка на медиа-сообщение, на которое пользователь ответил']
column_counter = 0
for one_header in list_of_headers :
 excel_worksheet.write(0,column_counter,one_header,font_for_headers)
 column_counter += 1
 
row_counter = 1

# Построчно заполняем Excel-файл
for string_from_query in tuple_from_query:
 for column_counter in range (12) :
  ogg_way = ''
  mp3_way = ''
  image_way = ''
  results_image_field = ''
  results_audio_field = ''
  audio_message = ''
  
  if column_counter == 5 :
   not_converted_time = string_from_query[column_counter]
   converted_time = datetime.datetime.utcfromtimestamp(int(str(not_converted_time)[0:10]))
   excel_worksheet.write(row_counter,column_counter,str(converted_time))
   
  elif column_counter in {10, 11} :
# Осуществляем поиск в строках полей "attach", "nested" подстрок со ссылками на голосовые сообщения, либо на изображения и записываем результат в Excel-файл построчно
   results_image_field = re.findall(r'(https.{211,216}album)',str(string_from_query[column_counter]))
   audio_message = re.findall(r'(https.*\.ogg)|(https.*\.mp3)',str(string_from_query[column_counter]))
   if len(results_image_field) != 0 :
    i = 0
    while i < len(results_image_field) :
     image_way = image_way + results_image_field[i]
     image_way = image_way + '\n'
     i += 1 
    excel_worksheet.write(row_counter,column_counter,str(image_way))
   elif len(audio_message) != 0 :
    for ogg_way, mp3_way in audio_message :
     if ogg_way != '':
      results_audio_field = results_audio_field + ogg_way
      results_audio_field = results_audio_field + '\n'
     elif mp3_way != '':
      results_audio_field = results_audio_field + mp3_way
    excel_worksheet.write(row_counter,column_counter,results_audio_field)
   else :
    excel_worksheet.write(row_counter,column_counter,'')

  else :
   excel_worksheet.write(row_counter,column_counter,str(string_from_query[column_counter]))
   
 row_counter += 1

excel_file.close()
