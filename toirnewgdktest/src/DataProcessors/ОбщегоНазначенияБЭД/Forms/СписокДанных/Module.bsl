#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Инициализировать();

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовФормы

&НаКлиенте
Процедура СписокДанныхВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	ТекущиеДанные = Элемент.ТекущиеДанные;
	
	Если ТекущиеДанные <> Неопределено Тогда
		ПоказатьЗначение(, ТекущиеДанные.Значение);
	КонецЕсли;

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура Инициализировать()
	
	Если ЗначениеЗаполнено(Параметры.Заголовок) Тогда
		Заголовок = Параметры.Заголовок;
		ОтображатьЗаголовок = Истина;
	КонецЕсли;
	
	Для Каждого ЭлементСписок Из Параметры.СписокДанных Цикл
		СписокДанных.Добавить(ЭлементСписок.Значение, ЭлементСписок.Представление);
	КонецЦикла;
	
КонецПроцедуры

#КонецОбласти