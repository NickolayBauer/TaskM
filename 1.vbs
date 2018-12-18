Dim Msg, Style, Title, Response, x, p, y 				'Объявляем необходимые переменные

Title = "Савин Михаил: Задание #1" 						'Присваем переменной Title имя заголовка программы
Style = vbOKOnly 										'Присваем переменной Style кнопку "ОК"

x = inputBox("Введите x кг", Title) 					'Вводим переменную x
p = inputBox("Введите цену p", Title)					'Вводим переменную p
y = inputBox("Введите количество p кг", Title)  		'Вводим переменную y

if isNumeric(x) and isNumeric(p) and isNumeric(y) Then 'Проверяем, являются ли вводимые значения цифрами
	if  x >  0 and p  >= 0 and y > 0 Then
													   'Проверяем, являются ли переменные положительными 
		Style = Style + vbInformation				   'Добавляем к переменной Style информационную пиктограмму
													   'Считаем результат и присваиваем его переменной Msg
		Msg = "Стоимость "&x&" кг товара = "&Round(p/x * y, 2)&" RUB"
		Response = MsgBox(Msg, Style, Title)		   'Вызываем MsgBox
    
	Else                                               'Если одна или несколько переменные не являются положительными 
												       'Также цена может быть равно нулю но не может быть отрицательной
		Style = Style + vbCritical					   'Добавляем переменной Style пиктограмму ошибки
													   'Вызываем MsgBox с сообщением об ошибке 
		Msg = "Была допущена ошибка  ввода(Одно или несколько значений не являются положительными либо цена отрицательная) попробуйте снова"
		Response = MsgBox(Msg, Style, Title)

	End if 
Else												   'Если одна или несколько переменные не являются цифрами
	Style = Style + vbCritical						   'Добавляем переменной Style пиктограмму ошибки
													   'Вызываем MsgBox с сообщением об ошибке 
	Msg = "Была допущена ошибка ввода (Одно или несколько значений не являются цифрами) попробуйте снова"
	Response = MsgBox(Msg, Style, Title)
	
End if
