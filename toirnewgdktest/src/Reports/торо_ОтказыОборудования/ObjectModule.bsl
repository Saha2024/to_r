#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ОбработчикиСобытий
Процедура ПриКомпоновкеРезультата(ДокументРезультат, ДанныеРасшифровки, СтандартнаяОбработка)
	
	ТекИерархия = торо_ОтчетыСервер.ПолучитьЗначениеСтруктурыИерархии(КомпоновщикНастроек);
	
	торо_ОтчетыСервер.УстановитьЗапросыНаборовДанныхИерархииОР(СхемаКомпоновкиДанных, ТекИерархия, "ДатаКон");
	
	Если ТекИерархия.СтроитсяАвтоматически Тогда
				
		СхемаКомпоновкиДанных.НаборыДанных.Объекты.Запрос =
		ПолучитьСодержательнуюЧастьЗапроса()+
		"ВЫБРАТЬ
		|	ВТ_Данные.Организация КАК Организация,
		|	ВТ_Данные.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_Данные.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
		|	ВТ_Данные.ДефектПричина КАК ДефектПричина,
		|	ВТ_Данные.ВидДефекта КАК ВидДефекта,
		|	ВТ_Данные.Регистратор КАК Регистратор,
		|	ВТ_Данные.ОтказавшийОбъектРемонта КАК ОтказавшийОбъектРемонта,
		|	ВТ_Данные.Количество КАК Количество,
		|	ВТ_Данные.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
		|	ВТ_Данные.ДлительностьРемонта КАК ДлительностьРемонта,
		|	ВТ_Данные.ДатаДефекта КАК ДатаДефекта,
		|	ВТ_Данные.МатериальныеЗатраты КАК МатериальныеЗатраты,
		|	ВТ_Данные.ТрудовыеЗатраты КАК ТрудовыеЗатраты,
		|	ВТ_Данные.ОбъектРемонта." + ТекИерархия.РеквизитОР + " КАК ОбъектИерархии,
		|	ВТ_Данные.ГарантийныйРемонт КАК ГарантийныйРемонт
		|ИЗ
		|	ВТ_Данные КАК ВТ_Данные";
		
		торо_ОтчетыКлиентСервер.УстановитьТипИерархическойГруппировкиВНастройках(КомпоновщикНастроек, "ОбъектИерархии", ТипГруппировкиКомпоновкиДанных.Иерархия);
		
	Иначе

		торо_ОтчетыКлиентСервер.УстановитьТипИерархическойГруппировкиВНастройках(КомпоновщикНастроек, "ОбъектИерархии", ТипГруппировкиКомпоновкиДанных.ТолькоИерархия);

		СхемаКомпоновкиДанных.НаборыДанных.Объекты.Запрос = 
		ПолучитьСодержательнуюЧастьЗапроса()+
		"ВЫБРАТЬ
		|	ВТ_Данные.Организация КАК Организация,
		|	ВТ_Данные.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_Данные.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
		|	ВТ_Данные.ДефектПричина КАК ДефектПричина,
		|	ВТ_Данные.ВидДефекта КАК ВидДефекта,
		|	ВТ_Данные.Регистратор КАК Регистратор,
		|	ВТ_Данные.ОтказавшийОбъектРемонта КАК ОтказавшийОбъектРемонта,
		|	ВТ_Данные.Количество КАК Количество,
		|	ВТ_Данные.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
		|	ВТ_Данные.ДлительностьРемонта КАК ДлительностьРемонта,
		|	ВТ_Данные.ДатаДефекта КАК ДатаДефекта,
		|	ВТ_Данные.МатериальныеЗатраты КАК МатериальныеЗатраты,
		|	ВТ_Данные.ТрудовыеЗатраты КАК ТрудовыеЗатраты,
		|	ВТ_Данные.ОбъектРемонта КАК ОбъектИерархии,
		|	ВТ_Данные.ГарантийныйРемонт КАК ГарантийныйРемонт
		|ИЗ
		|	ВТ_Данные КАК ВТ_Данные";
		
	КонецЕсли;
	
	НастроитьДополнениеПоДатеДефекта(КомпоновщикНастроек.Настройки);
	
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

Процедура НастроитьДополнениеПоДатеДефекта(НастройкиСхемы)
	
	ПолеГруппировки = ПолучитьПолеГруппировкиКомпоновкиДанныхДляНастройкиДополнения(НастройкиСхемы, "ДатаДефекта");
	Если ПолеГруппировки <> Неопределено Тогда
		УстановитьТипДополненияПериодаДляПоляГруппировки(ПолеГруппировки, НастройкиСхемы);
		УстановитьДатыНачалаИОкончанияДополнения(ПолеГруппировки, НастройкиСхемы);
	КонецЕсли;
	
КонецПроцедуры

Функция ПолучитьПолеГруппировкиКомпоновкиДанныхДляНастройкиДополнения(НастройкиСхемы, ИмяГруппировки)
	
	Для Каждого ЭлементНастройки Из НастройкиСхемы.Структура Цикл
		Если ТипЗнч(ЭлементНастройки) = Тип("ТаблицаКомпоновкиДанных") Тогда
			Для Каждого Колонка Из ЭлементНастройки.Колонки Цикл
				Если Колонка.Имя = ИмяГруппировки Тогда
					Возврат Колонка.ПоляГруппировки.Элементы[0];
				КонецЕсли;
			КонецЦикла;
		КонецЕсли;
	КонецЦикла;
	
КонецФункции

Процедура УстановитьТипДополненияПериодаДляПоляГруппировки(ПолеГруппировкиКомпоновкиДанных, НастройкиСхемы)

	ПараметрПериодичностьОтчета = НастройкиСхемы.ПараметрыДанных.Элементы.Найти("Периодичность");
	СтрокаПараметр = КомпоновщикНастроек.ПользовательскиеНастройки.Элементы.Найти(ПараметрПериодичностьОтчета.ИдентификаторПользовательскойНастройки);
	
	Если СтрокаПараметр <> Неопределено Тогда
		ТипДополненияПериодаПеречисление = СтрокаПараметр.Значение;
		Если ТипДополненияПериодаПеречисление = Перечисления.Периодичность.Месяц Тогда
			ПолеГруппировкиКомпоновкиДанных.ТипДополнения = ТипДополненияПериодаКомпоновкиДанных.Месяц;
		ИначеЕсли ТипДополненияПериодаПеречисление = Перечисления.Периодичность.День Тогда
			ПолеГруппировкиКомпоновкиДанных.ТипДополнения = ТипДополненияПериодаКомпоновкиДанных.День;
		ИначеЕсли ТипДополненияПериодаПеречисление = Перечисления.Периодичность.Год Тогда
			ПолеГруппировкиКомпоновкиДанных.ТипДополнения = ТипДополненияПериодаКомпоновкиДанных.Год;
		ИначеЕсли ТипДополненияПериодаПеречисление = Перечисления.Периодичность.Квартал Тогда
			ПолеГруппировкиКомпоновкиДанных.ТипДополнения = ТипДополненияПериодаКомпоновкиДанных.Квартал;
		ИначеЕсли ТипДополненияПериодаПеречисление = Перечисления.Периодичность.Неделя Тогда
			ПолеГруппировкиКомпоновкиДанных.ТипДополнения = ТипДополненияПериодаКомпоновкиДанных.Неделя;
		КонецЕсли;
	КонецЕсли;
	
КонецПроцедуры

Процедура УстановитьДатыНачалаИОкончанияДополнения(ПолеГруппировкиКомпоновкиДанных, НастройкиСхемы)

	ПараметрПериодОтчета = НастройкиСхемы.ПараметрыДанных.Элементы.Найти("Период");
	СтрокаПараметр = КомпоновщикНастроек.ПользовательскиеНастройки.Элементы.Найти(ПараметрПериодОтчета.ИдентификаторПользовательскойНастройки);
	
	Если СтрокаПараметр <> Неопределено И СтрокаПараметр.Использование Тогда
		ПолеГруппировкиКомпоновкиДанных.НачалоПериода = СтрокаПараметр.Значение.ДатаНачала;
		ПолеГруппировкиКомпоновкиДанных.КонецПериода = СтрокаПараметр.Значение.ДатаОкончания;
	КонецЕсли;

КонецПроцедуры

Функция ПолучитьСодержательнуюЧастьЗапроса() 
	
	ЗапросТекст =
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ РАЗЛИЧНЫЕ
	|	торо_ВыявленныеДефекты.ID КАК ID
	|ПОМЕСТИТЬ IDРемонтов
	|ИЗ
	|	РегистрСведений.торо_ВыявленныеДефекты КАК торо_ВыявленныеДефекты
	|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_ОтмененныеРемонты КАК торо_ОтмененныеРемонты
	|		ПО торо_ВыявленныеДефекты.ID = торо_ОтмененныеРемонты.ID
	|ГДЕ
	|	торо_ОтмененныеРемонты.Регистратор ЕСТЬ NULL
	|	И торо_ВыявленныеДефекты.Период МЕЖДУ &ДатаНач И &ДатаКон
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ID
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	IDРемонтов.ID КАК ID,
	|	МАКСИМУМ(ЦеныНоменклатуры.Период) КАК Период,
	|	ЦеныНоменклатуры.Номенклатура КАК Номенклатура,
	|	ЦеныНоменклатуры.ВидЦены КАК ВидЦены,
	|	ЦеныНоменклатуры.Характеристика КАК Характеристика
	|ПОМЕСТИТЬ ПериодыЦенНоменклатуры
	|ИЗ
	|	IDРемонтов КАК IDРемонтов
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрНакопления.торо_ФактическиеМатериальныеЗатратыРемонтныхРабот КАК торо_ФактическиеМатериальныеЗатратыРемонтныхРабот
	|			ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.ЦеныНоменклатуры КАК ЦеныНоменклатуры
	|				ЛЕВОЕ СОЕДИНЕНИЕ Константа.торо_ТипЦеныДляРасчетаСебестоимостиРемонта КАК торо_ТипЦеныДляРасчетаСебестоимостиРемонта
	|				ПО (ИСТИНА)
	|			ПО торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Номенклатура = ЦеныНоменклатуры.Номенклатура
	|				И торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.ХарактеристикаНоменклатуры = ЦеныНоменклатуры.Характеристика
	|				И торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Период >= ЦеныНоменклатуры.Период
	|				И (ЦеныНоменклатуры.ВидЦены = торо_ТипЦеныДляРасчетаСебестоимостиРемонта.Значение)
	|		ПО IDРемонтов.ID = торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.РемонтыОборудования_ID
	|
	|СГРУППИРОВАТЬ ПО
	|	IDРемонтов.ID,
	|	ЦеныНоменклатуры.Номенклатура,
	|	ЦеныНоменклатуры.ВидЦены,
	|	ЦеныНоменклатуры.Характеристика
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Период,
	|	Номенклатура,
	|	Характеристика,
	|	ВидЦены,
	|	ID
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	IDРемонтов.ID КАК ID,
	|	торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Номенклатура КАК Номенклатура,
	|	торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.ХарактеристикаНоменклатуры КАК ХарактеристикаНоменклатуры,
	|	СУММА(торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Количество) КАК Количество,
	|	СУММА(торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Количество * ЕСТЬNULL(ЦеныНоменклатуры.Цена, 0)) КАК Сумма
	|ПОМЕСТИТЬ МатериальныеЗатраты
	|ИЗ
	|	РегистрНакопления.торо_ФактическиеМатериальныеЗатратыРемонтныхРабот КАК торо_ФактическиеМатериальныеЗатратыРемонтныхРабот
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ IDРемонтов КАК IDРемонтов
	|		ПО (IDРемонтов.ID = торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.РемонтыОборудования_ID)
	|		ЛЕВОЕ СОЕДИНЕНИЕ ПериодыЦенНоменклатуры КАК ПериодыЦенНоменклатуры
	|			ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.ЦеныНоменклатуры КАК ЦеныНоменклатуры
	|			ПО ПериодыЦенНоменклатуры.Период = ЦеныНоменклатуры.Период
	|				И ПериодыЦенНоменклатуры.Номенклатура = ЦеныНоменклатуры.Номенклатура
	|				И ПериодыЦенНоменклатуры.Характеристика = ЦеныНоменклатуры.Характеристика
	|				И ПериодыЦенНоменклатуры.ВидЦены = ЦеныНоменклатуры.ВидЦены
	|		ПО торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Номенклатура = ПериодыЦенНоменклатуры.Номенклатура
	|			И торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.ХарактеристикаНоменклатуры = ПериодыЦенНоменклатуры.Характеристика
	|			И торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.РемонтыОборудования_ID = ПериодыЦенНоменклатуры.ID
	|
	|СГРУППИРОВАТЬ ПО
	|	IDРемонтов.ID,
	|	торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.Номенклатура,
	|	торо_ФактическиеМатериальныеЗатратыРемонтныхРабот.ХарактеристикаНоменклатуры
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ID
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	IDРемонтов.ID КАК ID,
	|	СУММА(торо_ФактическиеТрудовыеЗатратыРемонтныхРабот.Сумма) КАК Сумма
	|ПОМЕСТИТЬ ТрудовыеЗатраты
	|ИЗ
	|	IDРемонтов КАК IDРемонтов
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрНакопления.торо_ФактическиеТрудовыеЗатратыРемонтныхРабот КАК торо_ФактическиеТрудовыеЗатратыРемонтныхРабот
	|		ПО IDРемонтов.ID = торо_ФактическиеТрудовыеЗатратыРемонтныхРабот.РемонтыОборудования_ID
	|
	|СГРУППИРОВАТЬ ПО
	|	IDРемонтов.ID
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ID
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	ВЫБОР 
	|		КОГДА торо_ОбъектыРемонта.ВнешнийОбъект ТОГДА торо_ОбъектыРемонта.Контрагент
	|		ИНАЧЕ торо_ОбъектыРемонта.Организация
	|	КОНЕЦ КАК Организация,
	|	торо_ВыявленныеДефекты.ОбъектРемонта КАК ОбъектРемонта,
	|	торо_ВыявленныеДефекты.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
	|	торо_ВыявленныеДефекты.ДефектПричина КАК ДефектПричина,
	|	торо_ВыявленныеДефекты.ВидДефекта КАК ВидДефекта,
	|	торо_ВыявленныеДефекты.Регистратор КАК Регистратор,
	|	торо_ВыявленныеДефекты.ОбъектРемонта КАК ОтказавшийОбъектРемонта,
	|	торо_ВыявленныеДефекты.ID КАК ID,
	|	торо_ВыявленныеДефекты.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
	|	СУММА(ЕСТЬNULL(РАЗНОСТЬДАТ(торо_СвернутыеФактическиеДатыРемонтов.ДатаНачала, торо_СвернутыеФактическиеДатыРемонтов.ДатаОкончания, СЕКУНДА), 0)) КАК ДлительностьРемонта,
	|	ВЫБОР &Периодичность
	|		КОГДА ЗНАЧЕНИЕ(Перечисление.Периодичность.День)
	|			ТОГДА НАЧАЛОПЕРИОДА(торо_ВыявленныеДефекты.Период, ДЕНЬ)
	|		КОГДА ЗНАЧЕНИЕ(Перечисление.Периодичность.Месяц)
	|			ТОГДА НАЧАЛОПЕРИОДА(торо_ВыявленныеДефекты.Период, МЕСЯЦ)
	|		КОГДА ЗНАЧЕНИЕ(Перечисление.Периодичность.Год)
	|			ТОГДА НАЧАЛОПЕРИОДА(торо_ВыявленныеДефекты.Период, ГОД)
	|		КОГДА ЗНАЧЕНИЕ(Перечисление.Периодичность.Квартал)
	|			ТОГДА НАЧАЛОПЕРИОДА(торо_ВыявленныеДефекты.Период, КВАРТАЛ)
	|		КОГДА ЗНАЧЕНИЕ(Перечисление.Периодичность.Неделя)
	|			ТОГДА НАЧАЛОПЕРИОДА(торо_ВыявленныеДефекты.Период, НЕДЕЛЯ)
	|	КОНЕЦ КАК ДатаДефекта
	|ПОМЕСТИТЬ БезМатИТрудовых
	|ИЗ
	|	РегистрСведений.торо_ВыявленныеДефекты КАК торо_ВыявленныеДефекты
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ IDРемонтов КАК IDРемонтов
	|			ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_СвернутыеФактическиеДатыРемонтов КАК торо_СвернутыеФактическиеДатыРемонтов
	|			ПО IDРемонтов.ID = торо_СвернутыеФактическиеДатыРемонтов.IDРемонта
	|		ПО торо_ВыявленныеДефекты.ID = IDРемонтов.ID
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.торо_ОбъектыРемонта КАК торо_ОбъектыРемонта
	|		ПО торо_ВыявленныеДефекты.ОбъектРемонта = торо_ОбъектыРемонта.Ссылка
	|
	|СГРУППИРОВАТЬ ПО
	|	ВЫБОР 
	|		КОГДА торо_ОбъектыРемонта.ВнешнийОбъект ТОГДА торо_ОбъектыРемонта.Контрагент
	|		ИНАЧЕ торо_ОбъектыРемонта.Организация
	|	КОНЕЦ,
	|	торо_ВыявленныеДефекты.ОбъектРемонта,
	|	торо_ВыявленныеДефекты.ОтказавшийЭлемент,
	|	торо_ВыявленныеДефекты.ДефектПричина,
	|	торо_ВыявленныеДефекты.Регистратор,
	|	торо_ВыявленныеДефекты.ВидДефекта,
	|	торо_ВыявленныеДефекты.ДефектВыявленВПроцессе,
	|	торо_ВыявленныеДефекты.Период,
	|	торо_ВыявленныеДефекты.ID,
	|	торо_ВыявленныеДефекты.ОбъектРемонта
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	ID
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	БезМатИТрудовых.Организация КАК Организация,
	|	БезМатИТрудовых.ОбъектРемонта КАК ОбъектРемонта,
	|	БезМатИТрудовых.ОтказавшийЭлемент КАК ОтказавшийЭлемент,
	|	БезМатИТрудовых.ДефектПричина КАК ДефектПричина,
	|	БезМатИТрудовых.ВидДефекта КАК ВидДефекта,
	|	БезМатИТрудовых.Регистратор КАК Регистратор,
	|	БезМатИТрудовых.ОтказавшийОбъектРемонта КАК ОтказавшийОбъектРемонта,
	|	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ БезМатИТрудовых.ID) КАК Количество,
	|	БезМатИТрудовых.ДефектВыявленВПроцессе КАК ДефектВыявленВПроцессе,
	|	СУММА(БезМатИТрудовых.ДлительностьРемонта) КАК ДлительностьРемонта,
	|	БезМатИТрудовых.ДатаДефекта КАК ДатаДефекта,
	|	СУММА(ЕСТЬNULL(МатериальныеЗатраты.Сумма, 0)) КАК МатериальныеЗатраты,
	|	СУММА(ЕСТЬNULL(ТрудовыеЗатраты.Сумма, 0)) КАК ТрудовыеЗатраты,
	|	МАКСИМУМ(торо_ВыявленныеДефектыСписокДефектов.ГарантийныйРемонт) КАК ГарантийныйРемонт
	|ПОМЕСТИТЬ ВТ_Данные
	|ИЗ
	|	БезМатИТрудовых КАК БезМатИТрудовых
	|		ЛЕВОЕ СОЕДИНЕНИЕ МатериальныеЗатраты КАК МатериальныеЗатраты
	|		ПО БезМатИТрудовых.ID = МатериальныеЗатраты.ID
	|		ЛЕВОЕ СОЕДИНЕНИЕ ТрудовыеЗатраты КАК ТрудовыеЗатраты
	|		ПО БезМатИТрудовых.ID = ТрудовыеЗатраты.ID
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.торо_ВыявленныеДефекты.СписокДефектов КАК торо_ВыявленныеДефектыСписокДефектов
	|		ПО БезМатИТрудовых.ID = торо_ВыявленныеДефектыСписокДефектов.ID
	|
	|СГРУППИРОВАТЬ ПО
	|	БезМатИТрудовых.Организация,
	|	БезМатИТрудовых.ОбъектРемонта,
	|	БезМатИТрудовых.ОтказавшийЭлемент,
	|	БезМатИТрудовых.ДефектПричина,
	|	БезМатИТрудовых.ВидДефекта,
	|	БезМатИТрудовых.Регистратор,
	|	БезМатИТрудовых.ОтказавшийОбъектРемонта,
	|	БезМатИТрудовых.ДефектВыявленВПроцессе,
	|	БезМатИТрудовых.ДатаДефекта"
		
	+ "
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|";
	
	Возврат ЗапросТекст;
	
КонецФункции

#КонецОбласти

#КонецЕсли