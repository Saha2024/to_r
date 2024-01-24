
#Область ОбработчикиСобытийФормы
&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	ОбщегоНазначенияКлиентСервер.УстановитьПараметрДинамическогоСписка(СписокАктов, "Ссылка", Параметры.ОбъектОтбора);
КонецПроцедуры
#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыСписокАктов
&НаКлиенте
Процедура СписокАктовВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	ПоказатьЗначение(Неопределено, Элементы.СписокАктов.ТекущиеДанные.ДокументТОИР);
КонецПроцедуры
#КонецОбласти