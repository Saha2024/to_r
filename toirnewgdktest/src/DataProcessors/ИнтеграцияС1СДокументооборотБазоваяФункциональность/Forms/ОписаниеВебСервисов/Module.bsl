#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	ОписаниеВебСервисов = ИнтеграцияС1СДокументооборотБазоваяФункциональность.ОписаниеВебСервисов();
	
	Параметры.Свойство("Ссылка", Ссылка);
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	Если ЗначениеЗаполнено(Ссылка) Тогда
		ПодключитьОбработчикОжидания("ПерейтиКСсылке", 0.2, Истина);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаКлиенте
Процедура ПерейтиКСсылке()
	
	// При проблемах с совместимостью браузеров откроем документ без позиционирования
	// на ссылке, проглотив исключение как некритичное.
	Попытка
		Документ = Элементы.ПолеHTMLДокумента.Документ;
		Элемент = Документ.getElementByID(Ссылка);
		Если Элемент <> Неопределено Тогда
			Элемент.scrollIntoView(Истина)
		КонецЕсли;
	Исключение
		Возврат;
	КонецПопытки;
	
КонецПроцедуры

#КонецОбласти