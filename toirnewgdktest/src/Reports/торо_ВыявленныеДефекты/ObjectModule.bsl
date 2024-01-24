#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда
	
#Область ОбработчикиСобытий

Процедура ПриКомпоновкеРезультата(ДокументРезультат, ДанныеРасшифровки, СтандартнаяОбработка)
	
	ТекИерархия = торо_ОтчетыСервер.ПолучитьЗначениеСтруктурыИерархии(КомпоновщикНастроек);
	
	торо_ОтчетыСервер.УстановитьЗапросыНаборовДанныхИерархииОР(СхемаКомпоновкиДанных, ТекИерархия, "ДатаКон");
	
	Если ТекИерархия.СтроитсяАвтоматически Тогда
				
		СхемаКомпоновкиДанных.НаборыДанных.Объекты.Запрос =
		ПолучитьСодержательнуюЧастьЗапроса()+
		"ВЫБРАТЬ
		|	ВТ_Данные.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_Данные.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
		|	ВТ_Данные.ДефектПричина КАК ДефектПричина,
		|	ВТ_Данные.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
		|	ВТ_Данные.ВидДефекта КАК ВидДефекта,
		|	ВТ_Данные.ОформленнаяЗаявка КАК ОформленнаяЗаявка,
		|	ВТ_Данные.Документ КАК Документ,
		|	ВТ_Данные.Дата КАК Дата,
		|	ВТ_Данные.НеисправныйОбъектРемонта КАК НеисправныйОбъектРемонта,
		|	ВТ_Данные.ID КАК ID,
		|	ВТ_Данные.ОписаниеДефекта КАК ОписаниеДефекта,
		|	ВТ_Данные.КритичностьДефекта КАК КритичностьДефекта,
		|	ВТ_Данные.Закрыт КАК Закрыт,
		|	ВТ_Данные.ПричинаЗакрытия КАК ПричинаЗакрытия,
		|	ВТ_Данные.СтатусРемонта КАК СтатусРемонта,
		|	ВТ_Данные.ПоследствияДефекта КАК ПоследствияДефекта,
		|	ВТ_Данные.НаправлениеНаУстранение КАК НаправлениеНаУстранение,
		|	ВТ_Данные.ОбъектРемонта." + ТекИерархия.РеквизитОР + " КАК ОбъектИерархии,
		|	ВТ_Данные.ГарантийныйРемонт КАК ГарантийныйРемонт,
		|	ВТ_Данные.ТиповойДефект КАК ТиповойДефект
		|ИЗ
		|	ВТ_Данные КАК ВТ_Данные";
		
		торо_ОтчетыКлиентСервер.УстановитьТипИерархическойГруппировкиВНастройках(КомпоновщикНастроек, "ОбъектИерархии", ТипГруппировкиКомпоновкиДанных.Иерархия);
		
	Иначе

		торо_ОтчетыКлиентСервер.УстановитьТипИерархическойГруппировкиВНастройках(КомпоновщикНастроек, "ОбъектИерархии", ТипГруппировкиКомпоновкиДанных.ТолькоИерархия);

		СхемаКомпоновкиДанных.НаборыДанных.Объекты.Запрос = 
		ПолучитьСодержательнуюЧастьЗапроса()+
		"ВЫБРАТЬ
		|	ВТ_Данные.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_Данные.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
		|	ВТ_Данные.ДефектПричина КАК ДефектПричина,
		|	ВТ_Данные.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
		|	ВТ_Данные.ВидДефекта КАК ВидДефекта,
		|	ВТ_Данные.ОформленнаяЗаявка КАК ОформленнаяЗаявка,
		|	ВТ_Данные.Документ КАК Документ,
		|	ВТ_Данные.Дата КАК Дата,
		|	ВТ_Данные.НеисправныйОбъектРемонта КАК НеисправныйОбъектРемонта,
		|	ВТ_Данные.ID КАК ID,
		|	ВТ_Данные.ОписаниеДефекта КАК ОписаниеДефекта,
		|	ВТ_Данные.КритичностьДефекта КАК КритичностьДефекта,
		|	ВТ_Данные.Закрыт КАК Закрыт,
		|	ВТ_Данные.ПричинаЗакрытия КАК ПричинаЗакрытия,
		|	ВТ_Данные.СтатусРемонта КАК СтатусРемонта,
		|	ВТ_Данные.ПоследствияДефекта КАК ПоследствияДефекта,
		|	ВТ_Данные.НаправлениеНаУстранение КАК НаправлениеНаУстранение,
		|	ВТ_Данные.ОбъектРемонта КАК ОбъектИерархии,
		|	ВТ_Данные.ГарантийныйРемонт КАК ГарантийныйРемонт,
		|	ВТ_Данные.ТиповойДефект КАК ТиповойДефект
		|ИЗ
		|	ВТ_Данные КАК ВТ_Данные";
		
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ОпределитьНастройкиФормы(Форма, КлючВарианта, Настройки) Экспорт
	
	Настройки.ЗагрузитьНастройкиПриИзмененииПараметров = ЗагрузитьНастройкиПриИзмененииПараметров();

КонецПроцедуры

Функция ЗагрузитьНастройкиПриИзмененииПараметров()  
	
	Параметры = Новый Массив;
	Параметры.Добавить(Новый ПараметрКомпоновкиДанных("ИерархияТип"));	
	Возврат Параметры;
	
КонецФункции

Функция ПолучитьСодержательнуюЧастьЗапроса() 
	
	ЗапросТекст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ
	              |	торо_ВыявленныеДефекты.ОбъектРемонта КАК ОбъектРемонта,
	              |	торо_ВыявленныеДефекты.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
	              |	торо_ВыявленныеДефекты.ДефектПричина КАК ДефектПричина,
	              |	торо_ВыявленныеДефекты.Регистратор.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
	              |	торо_ВыявленныеДефекты.ВидДефекта КАК ВидДефекта,
	              |	торо_ВыявленныеДефекты.Регистратор КАК Документ,
	              |	торо_ВыявленныеДефекты.Период КАК Дата,
	              |	торо_ВыявленныеДефекты.ОбъектРемонта КАК НеисправныйОбъектРемонта,
	              |	торо_ВыявленныеДефекты.ID КАК ID,
	              |	торо_ВыявленныеДефекты.ДефектОписание КАК ОписаниеДефекта,
	              |	торо_ВыявленныеДефекты.КритичностьДефекта КАК КритичностьДефекта,
	              |	торо_ВыявленныеДефекты.ПоследствияДефекта КАК ПоследствияДефекта,
	              |	торо_ВыявленныеДефекты.НаправлениеНаУстранение КАК НаправлениеНаУстранение,
	              |	ЛОЖЬ КАК ГарантийныйРемонт,
	              |	торо_ВыявленныеДефекты.ТиповойДефект КАК ТиповойДефект
	              |ПОМЕСТИТЬ ВТ_ДанныеНачальная
	              |ИЗ
	              |	РегистрСведений.торо_ВыявленныеДефекты КАК торо_ВыявленныеДефекты
	              |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_ОбщиеДанныеПоРемонтам КАК торо_ОбщиеДанныеПоРемонтам
	              |		ПО торо_ВыявленныеДефекты.ID = торо_ОбщиеДанныеПоРемонтам.IDРемонта
	              |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
	              |		ПО торо_ВыявленныеДефекты.ОтказавшийЭлемент = торо_ОбъектыРемонта.Ссылка
	              |{ГДЕ
	              |	(торо_ВыявленныеДефекты.Период МЕЖДУ &ДатаНач И &ДатаКон)}
	              |
	              |ИНДЕКСИРОВАТЬ ПО
	              |	ID
	              |;
	              |
	              |////////////////////////////////////////////////////////////////////////////////
	              |ВЫБРАТЬ РАЗРЕШЕННЫЕ
	              |	торо_СтатусыРемонтовСрезПоследних.IDРемонта КАК IDРемонта,
	              |	торо_СтатусыРемонтовСрезПоследних.Статус КАК Статус
	              |ПОМЕСТИТЬ ВТ_СтатусыДокументов
	              |ИЗ
	              |	РегистрСведений.торо_СтатусыРемонтов.СрезПоследних(
	              |			,
	              |			IDРемонта В
	              |				(ВЫБРАТЬ
	              |					ВТ_ДанныеНачальная.ID КАК ID
	              |				ИЗ
	              |					ВТ_ДанныеНачальная КАК ВТ_ДанныеНачальная)) КАК торо_СтатусыРемонтовСрезПоследних
	              |
	              |ИНДЕКСИРОВАТЬ ПО
	              |	IDРемонта
	              |;
	              |
	              |////////////////////////////////////////////////////////////////////////////////
	              |ВЫБРАТЬ РАЗРЕШЕННЫЕ
	              |	торо_ЗаявкиПоРемонтамСрезПоследних.IDРемонта КАК IDРемонта,
	              |	торо_ЗаявкиПоРемонтамСрезПоследних.ДокументЗаявка КАК ДокументЗаявка
	              |ПОМЕСТИТЬ ВТ_Заявки
	              |ИЗ
	              |	РегистрСведений.торо_ЗаявкиПоРемонтам.СрезПоследних(
	              |			,
	              |			IDРемонта В
	              |				(ВЫБРАТЬ
	              |					ВТ_ДанныеНачальная.ID
	              |				ИЗ
	              |					ВТ_ДанныеНачальная)) КАК торо_ЗаявкиПоРемонтамСрезПоследних
	              |
	              |ИНДЕКСИРОВАТЬ ПО
	              |	IDРемонта
	              |;
	              |
	              |////////////////////////////////////////////////////////////////////////////////
	              |ВЫБРАТЬ РАЗРЕШЕННЫЕ
	              |	ВТ_ДанныеНачальная.ОбъектРемонта КАК ОбъектРемонта,
	              |	ВТ_ДанныеНачальная.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
	              |	ВТ_ДанныеНачальная.ДефектПричина КАК ДефектПричина,
	              |	ВТ_ДанныеНачальная.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
	              |	ВТ_ДанныеНачальная.ВидДефекта КАК ВидДефекта,
	              |	ВЫБОР
	              |		КОГДА ВТ_Заявки.ДокументЗаявка ЕСТЬ NULL
	              |			ТОГДА ""Нет оформленных заявок""
	              |		ИНАЧЕ ВТ_Заявки.ДокументЗаявка
	              |	КОНЕЦ КАК ОформленнаяЗаявка,
	              |	ВТ_ДанныеНачальная.Документ КАК Документ,
	              |	ВТ_ДанныеНачальная.Дата КАК Дата,
	              |	ВТ_ДанныеНачальная.НеисправныйОбъектРемонта КАК НеисправныйОбъектРемонта,
	              |	ВТ_ДанныеНачальная.ID КАК ID,
	              |	торо_ОтмененныеРемонты.Отменен КАК Закрыт,
	              |	торо_ОтмененныеРемонты.ПричинаЗакрытия КАК ПричинаЗакрытия,
	              |	ВТ_ДанныеНачальная.ОписаниеДефекта КАК ОписаниеДефекта,
	              |	ВТ_ДанныеНачальная.КритичностьДефекта КАК КритичностьДефекта,
	              |	ВТ_ДанныеНачальная.ПоследствияДефекта КАК ПоследствияДефекта,
	              |	ВТ_ДанныеНачальная.НаправлениеНаУстранение КАК НаправлениеНаУстранение,
	              |	ВТ_ДанныеНачальная.ГарантийныйРемонт КАК ГарантийныйРемонт,
	              |	ВТ_ДанныеНачальная.ТиповойДефект КАК ТиповойДефект,
	              |	ВТ_СтатусыДокументов.Статус КАК СтатусРемонта
	              |ПОМЕСТИТЬ ВТ_Данные
	              |ИЗ
	              |	ВТ_ДанныеНачальная КАК ВТ_ДанныеНачальная
	              |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_Заявки КАК ВТ_Заявки
	              |		ПО ВТ_ДанныеНачальная.ID = ВТ_Заявки.IDРемонта
	              |		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_ОтмененныеРемонты КАК торо_ОтмененныеРемонты
	              |		ПО ВТ_ДанныеНачальная.ID = торо_ОтмененныеРемонты.ID
	              |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_СтатусыДокументов КАК ВТ_СтатусыДокументов
	              |		ПО ВТ_ДанныеНачальная.ID = ВТ_СтатусыДокументов.IDРемонта"
				  + " 
				  |; 
				  |  
				  |//////////////////////////////////////////////////////////////////////////////// 
				  |";
	
	Возврат ЗапросТекст;
	
КонецФункции

#КонецОбласти

#КонецЕсли