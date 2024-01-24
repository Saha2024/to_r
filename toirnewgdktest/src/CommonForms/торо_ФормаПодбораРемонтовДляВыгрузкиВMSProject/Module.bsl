
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	ПолучитьРемонты();
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ФильтрПриИзменении(Элемент)
	ПолучитьРемонты();
КонецПроцедуры

&НаКлиенте
Процедура ВыгрузитьВproject(Команда)
	
	ВыделенныеСтроки = Элементы.Ремонты.ВыделенныеСтроки;
	МассивID = Новый Массив;
	
	Для каждого Идентификатор Из ВыделенныеСтроки Цикл
		ДанныеСтроки = Ремонты.НайтиПоИдентификатору(Идентификатор);
		МассивID.Добавить(ДанныеСтроки.IDРемонта);
	КонецЦикла; 
	
	Если МассивID.Количество() > 0 Тогда
		Торо_ИнтеграцияСmsProject.ОтправитьВProject(МассивID, Неопределено, Истина);
	КонецЕсли; 

КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура ПолучитьРемонты()
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка КАК Заявка,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.НомерСтроки КАК НомерСтроки,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.ОбъектРемонта КАК Объект,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.ID КАК IDРемонта,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.ВидРемонтныхРабот КАК ВидРемонта,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаНачала,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаОкончания,
		|	торо_ЗаявкаНаРемонтРемонтыОборудования.ОписаниеРемонта КАК Описание
		|ИЗ
		|	Документ.торо_ЗаявкаНаРемонт.РемонтыОборудования КАК торо_ЗаявкаНаРемонтРемонтыОборудования
		|ГДЕ ИСТИНА" + Символы.ПС;
	
	Если ФлагПодразделения = ИСТИНА И Подразделения.Количество() <> 0 Тогда
		Запрос.Текст = Запрос.Текст+"И торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка.Подразделение В(&Подразделения)" + Символы.ПС;
	КонецЕсли;
	
	Если ФлагДатаЗаявкиОт И ЗначениеЗаполнено(ДатаЗаявкиОт) Тогда
		Запрос.Текст = Запрос.Текст+"И торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка.Дата >= &ДатаЗаявкиОт" + Символы.ПС;
	КонецЕсли;
	Если ФлагДатаЗаявкиДо И ЗначениеЗаполнено(ДатаЗаявкиДо) Тогда
		Запрос.Текст = Запрос.Текст+"И торо_ЗаявкаНаРемонтРемонтыОборудования.Ссылка.Дата <= &ДатаЗаявкиДо" + Символы.ПС;
	КонецЕсли;
	
	Если ФлагДатаРемонтаОт И ЗначениеЗаполнено(ДатаРемонтаОт) Тогда
		Запрос.Текст = Запрос.Текст+"И ((торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаНачала >= &ДатаРемонтаОт) ИЛИ (торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаОкончания >= &ДатаРемонтаОт))" + Символы.ПС;
	КонецЕсли;
	Если ФлагДатаРемонтаДо И ЗначениеЗаполнено(ДатаРемонтаДо) Тогда
		Запрос.Текст = Запрос.Текст+"И ((торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаНачала <= &ДатаРемонтаДо) ИЛИ (торо_ЗаявкаНаРемонтРемонтыОборудования.ДатаОкончания <= &ДатаРемонтаДо))" + Символы.ПС;
	КонецЕсли;
	
	Запрос.УстановитьПараметр("ДатаЗаявкиДо", ДатаЗаявкиДо);
	Запрос.УстановитьПараметр("ДатаЗаявкиОт", ДатаЗаявкиОт);
	Запрос.УстановитьПараметр("ДатаРемонтаДо", ДатаРемонтаДо);
	Запрос.УстановитьПараметр("ДатаРемонтаОт", ДатаРемонтаОт);
	Запрос.УстановитьПараметр("Подразделения", Подразделения);
	
	РезультатЗапроса = Запрос.Выполнить().Выгрузить();
	
	ЗначениеВРеквизитФормы(РезультатЗапроса, "Ремонты");
	
КонецПроцедуры

#КонецОбласти

