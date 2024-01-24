#Область ПрограммныйИнтерфейс

// Получение элемента справочника СтавкиНДС по перечислению СтавкиНДС
//
// Параметры:
//	ПеречислениеСтавкаНДС - ПеречислениеСсылка.СтавкиНДС - Значение ставки НДС.
//	ТипНалогообложенияНДС - ПеречислениеСсылка.ТипыНалогообложенияНДС - Налогообложение НДС, в рамках которого нужно найти ставку. 
//	                                                       Если Неопределено, то ставка получается по налогообложению ПродажаОблагаетсяНДС
// Возвращаемое значение:
//	СправочникСсылка.СтавкиНДС - Соответствующая перечислению ставка НДС.
//
Функция СтавкаНДСПоЗначениюПеречисления(ПеречислениеСтавкаНДС, ТипНалогообложенияНДС = Неопределено) Экспорт
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ ПЕРВЫЕ 1
		|	СтавкиНДС.Ссылка КАК СтавкаНДС
		|ИЗ
		|	Справочник.СтавкиНДС.ТипыНалогообложенияНДС КАК СтавкиНДС
		|ГДЕ
		|	СтавкиНДС.Ссылка.ПеречислениеСтавкаНДС = &ПеречислениеСтавкаНДС
		|	И СтавкиНДС.ТипНалогообложенияНДС = &ТипНалогообложенияНДС
		|
		|ОБЪЕДИНИТЬ ВСЕ
		|
		|ВЫБРАТЬ ПЕРВЫЕ 1
		|	СтавкиНДС.Ссылка КАК СтавкаНДС
		|ИЗ
		|	Справочник.СтавкиНДС КАК СтавкиНДС
		|ГДЕ
		|	СтавкиНДС.ПеречислениеСтавкаНДС = &ПеречислениеСтавкаНДС
		|	И &НалогообложениеНеУказано";
	
	Запрос.УстановитьПараметр("ПеречислениеСтавкаНДС", ПеречислениеСтавкаНДС);
	Если ТипНалогообложенияНДС = Неопределено Тогда
		Запрос.УстановитьПараметр("ТипНалогообложенияНДС", Перечисления.ТипыНалогообложенияНДС.ПродажаОблагаетсяНДС);
		Запрос.УстановитьПараметр("НалогообложениеНеУказано", Истина);
	Иначе
		Запрос.УстановитьПараметр("ТипНалогообложенияНДС", ТипНалогообложенияНДС);
		Запрос.УстановитьПараметр("НалогообложениеНеУказано", Ложь);
	КонецЕсли;
	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Если ВыборкаДетальныеЗаписи.Следующий() Тогда
		Возврат ВыборкаДетальныеЗаписи.СтавкаНДС
	КонецЕсли;
	
	Возврат Справочники.СтавкиНДС.ПустаяСсылка();
	
КонецФункции

// Функция возвращает процент НДС.
//
// Параметры:
//	СтавкаНДС - СправочникСсылка.СтавкиНДС - Значение ставки НДС.
//
// Возвращаемое значение:
//	Число - Процент НДС.
//
Функция ЗначениеСтавкиНДС(СтавкаНДС) Экспорт
	
	Если Не ЗначениеЗаполнено(СтавкаНДС) Тогда
		Возврат 0;
	КонецЕсли;
	
	Возврат ОбщегоНазначения.ЗначениеРеквизитаОбъекта(СтавкаНДС, "Ставка")
	
КонецФункции

#КонецОбласти
