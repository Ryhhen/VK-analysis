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
excel_worksheet.write(0,0,'local_id',font_for_headers)
excel_worksheet.write(0,1,'vk_id',font_for_headers)
excel_worksheet.write(0,2,'cnv_msg_id',font_for_headers)
excel_worksheet.write(0,3,'Название группы/канала',font_for_headers)
excel_worksheet.write(0,4,'ID собеседника',font_for_headers)
excel_worksheet.write(0,5,'Дата и время (UTC)',font_for_headers)
excel_worksheet.write(0,6,'Тип отправителя (1 - пользователь, 3 - группа/канал)',font_for_headers)
excel_worksheet.write(0,7,'ID отправителя',font_for_headers)
excel_worksheet.write(0,8,'Входящее ? (1 - да, 0 - нет)',font_for_headers)
excel_worksheet.write(0,9,'Текст сообщения',font_for_headers)
excel_worksheet.write(0,10,'Интернет-ссылка на медиа-сообщение',font_for_headers)
excel_worksheet.write(0,11,'Интернет-ссылка на медиа-сообщение, на которое пользователь ответил',font_for_headers)

row_counter = 1

# Построчно заполняем Excel-файл
for string_from_query in tuple_from_query:
 excel_worksheet.write(row_counter,0,str(string_from_query[0]))
 excel_worksheet.write(row_counter,1,str(string_from_query[1]))
 excel_worksheet.write(row_counter,2,str(string_from_query[2]))
 excel_worksheet.write(row_counter,3,str(string_from_query[3]))
 excel_worksheet.write(row_counter,4,str(string_from_query[4]))
 not_converted_time = string_from_query[5]
 converted_time = datetime.datetime.utcfromtimestamp(int(str(not_converted_time)[0:10]))
 excel_worksheet.write(row_counter,5,str(converted_time))
 excel_worksheet.write(row_counter,6,str(string_from_query[6]))
 excel_worksheet.write(row_counter,7,str(string_from_query[7]))
 excel_worksheet.write(row_counter,8,str(string_from_query[8]))
 excel_worksheet.write(row_counter,9,str(string_from_query[9]))
   
 results_image_field_attach = ''
 results_audio_field_attach = ''
 results_image_field_nested = ''
 results_audio_field_nested = ''
 ogg_way = ''
 mp3_way = ''
 image_way = ''

# Осуществляем поиск в строках полей "attach", "nested" подстрок со ссылками на голосовые сообщения, либо на изображения и записываем результат в Excel-файл построчно
 results_image_field_attach = re.findall(r'(https.{211,216}album)',str(string_from_query[10]))
 audio_message_attach = re.findall(r'(https.*\.ogg)|(https.*\.mp3)',str(string_from_query[10]))
 
 if len(results_image_field_attach) != 0 :
  i = 0
  while i < len(results_image_field_attach) :
   image_way = image_way + results_image_field_attach[i]
   image_way = image_way + '\n'
   i += 1 
  excel_worksheet.write(row_counter,10,str(image_way))
 elif len(audio_message_attach) != 0 :
  for ogg_way, mp3_way in audio_message_attach :
   if ogg_way != '':
    results_audio_field_attach = results_audio_field_attach + ogg_way
    results_audio_field_attach = results_audio_field_attach + '\n'
   elif mp3_way != '':
    results_audio_field_attach = results_audio_field_attach + mp3_way
  excel_worksheet.write(row_counter,10,results_audio_field_attach)
 else :
  excel_worksheet.write(row_counter,10,'')

 results_image_field_nested = re.findall(r'(https.{211,216}album)',str(string_from_query[11]))
 audio_message_nested = re.findall(r'(https.*\.ogg)|(https.*\.mp3)',str(string_from_query[11]))
 ogg_way = ''
 mp3_way = ''
 image_way = ''
 if len(results_image_field_nested) != 0 :
  i = 0
  while i < len(results_image_field_nested) :
   image_way = image_way + results_image_field_nested[i]
   image_way = image_way + '\n'
   i += 1 
  excel_worksheet.write(row_counter,11,str(image_way))
 elif len(audio_message_nested) != 0 :
  for ogg_way, mp3_way in audio_message_nested :
   if ogg_way != '':
    results_audio_field_nested = results_audio_field_nested + ogg_way
    results_audio_field_nested = results_audio_field_nested + '\n'
   elif mp3_way != '':
    results_audio_field_nested = results_audio_field_nested + mp3_way
  excel_worksheet.write(row_counter,11,results_audio_field_nested)
 else :
  excel_worksheet.write(row_counter,11,'')
 row_counter += 1
 
excel_file.close()