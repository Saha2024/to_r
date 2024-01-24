#Область ОбработчикиКомандФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)

	Если Параметры.Свойство("АвтоТест") Тогда // Возврат при получении формы для анализа.
		Возврат;
	КонецЕсли;
	
	РазрешитьРедактированиеИспользованиеХарактеристик = Истина;
	РазрешитьРедактированиеТипНоменклатуры = Истина;
	РазрешитьРедактированиеИспользованияСерий = Истина;
	
	СсылкаНаОбъект = Параметры.Ссылка;
	Элементы.ГруппаОбъектИспользуетсяСтраницы.ТекущаяСтраница = Элементы.ГруппаОбъектИспользуетсяНужноПроверить;

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийФормы

&НаКлиенте
Процедура РазрешитьРедактирование(Команда)

	Результат = Новый Массив();
	
	Если РазрешитьРедактированиеТипНоменклатуры Тогда
		Результат.Добавить("ТипНоменклатуры");
	КонецЕсли;
	
	Если РазрешитьРедактированиеИспользованиеХарактеристик Тогда
		Результат.Добавить("ИспользоватьХарактеристики");
		Результат.Добавить("ИспользованиеХарактеристик");
	КонецЕсли;
	
	Если РазрешитьРедактированиеИспользованияСерий Тогда
		Результат.Добавить("ИспользоватьСерии");
		Результат.Добавить("ИспользованиеСерий");
	КонецЕсли;
	
	Закрыть(Результат);

КонецПроцедуры

&НаКлиенте
Процедура ПроверитьИспользованиеОбъекта(Команда)
	
	ЕстьСсылкиНаОбъект = ОбщегоНазначенияВызовСервера.ЕстьСсылкиНаОбъект(СсылкаНаОбъект);
	Если ЕстьСсылкиНаОбъект Тогда
		Элементы.ГруппаОбъектИспользуетсяСтраницы.ТекущаяСтраница = Элементы.ГруппаОбъектИспользуетсяОбъектИспользуется;
		Элементы.ГруппаОбъектИспользуетсяНужноПроверить.Видимость = Ложь;
	Иначе
		Элементы.ГруппаОбъектИспользуетсяСтраницы.ТекущаяСтраница = Элементы.ГруппаОбъектИспользуетсяОбъектНеИспользуется;
		Элементы.ГруппаОбъектИспользуетсяНужноПроверить.Видимость = Ложь;
	КонецЕсли;

КонецПроцедуры

#КонецОбласти