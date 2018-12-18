Dim  Msg, Style, Title, Response, p, v     				'Объявляем необходимые переменные
    
Function funt_to_kg(p,v) 								'Объявляем функцию подсчёта объема

	p = p/2.205										    'Переводим в кг и кубические метры
	v = v/35.315
	funt_to_kg = Round(p/v,2)                           'Возвращаем результат

End Function

Title = "Савин Михаил: задание #3"					    'Присваием переменной Title имя заголовка программы
Style = vbOKOnly										'Присваием переменной Style кнопку "ОК"

p = InputBox("введите вес в фунтах",Title)				'Вводим вес
v = InputBox("ведите объем в кубических футах",Title)	'Вводим объем

if isNumeric(p) and isNumeric(v) Then					'Проверка на число
	if v > 0 and p >= 0 Then							'Проверка на положительное и на неотрицательное значение
		Msg = funt_to_kg(p,v)&" кг/м^3"					'Присваием переменной Msg результат выполнения функции funt_to_kg()
		Style = Style + vbInformation				    'Добавляем к переменной Style информационную пиктограмму
		Response = MsgBox(Msg, Style, Title)		    'Вызываем MsgBox
	Else
		Style = Style + vbCritical					    'Добавляем переменной Style пиктограмму ошибки
													    'Вызываем MsgBox с сообщением об ошибке 
		Msg = "Была допущена ошибка ввода (вес не может быть отрицательным, а объем не может быть отрицательным и нулевым) попробуйте снова"
		Response = MsgBox(Msg, Style, Title)
	End if
Else
	Style = Style + vbCritical						   'Добавляем переменной Style пиктограмму ошибки
													   'Вызываем MsgBox с сообщением об ошибке 
	Msg = "Была допущена ошибка ввода (Одно или несколько значений не являются цифрами) попробуйте снова"
	Response = MsgBox(Msg, Style, Title)	
End if


