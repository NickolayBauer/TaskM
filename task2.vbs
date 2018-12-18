Dim  Msg, Style, Title, Response, help,flag,error_flag, i, x(4)    				
														'Объявляем необходимые переменныеаммы
error_flag = True										'Считаем, что все числа буду числами
flag = True											  	'Считаем, что будущий массив ещё не отсортирован

Title = "Савин Михаил: Задание #2"					    'Присваем переменной Title имя заголовка программы
Style = vbOKOnly                                        'Присваем переменной Style кнопку "ОК"

For i = 0 To 3										    'Цикл для ввода элементов в массив
	Msg = "Ввведите элемент #"&(i+1)			        'Вызываем подсказку для ввода
	x(i) = InputBox(Msg, Title)					        'Вводим массив поэлементно
													    'Если элемент не является цифрой, то помещаем во флаг ошибки False
	if isNumeric(x(i)) = False Then
		error_flag  = False
	End if
Next

If error_flag Then      							    'Если все элементы массива являются числами
	Do While flag 	                                    'Пока массив не отсортирован, выполняем цикл сортировки
		flag = False									  
		For i = 0 to 2
			if x(i) > x(i+1) then 
				help = x(i)
				x(i) = x(i+1)
				x(i+1) = help
				flag = True
			End if 
		Next
	Loop
	
	Msg = Join(x, ", ")
	Style = Style + vbInformation				       'Добавляем к переменной Style информационную пиктограмму
	Response = MsgBox(Msg, Style, Title)
	
Else 
Style = Style + vbCritical						       'Добавляем переменной Style пиктограмму ошибки
Msg = "Была допущена ошибка ввода (Один из элементов массива не является числом)"
Response = MsgBox(Msg, Style, Title)
End if


