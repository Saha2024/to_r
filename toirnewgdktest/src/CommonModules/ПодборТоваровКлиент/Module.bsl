
#Область ПрограммныйИнтерфейс

#Область ГорячиеКлавиши

// Возвращает имя текущего элемента - строки поиска на форме подбора.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма подбора.
//
// Возвращаемое значение:
//		Строка - имя текущего элемента.
Функция ИмяТекущегоЭлементаСтрокиПоиска(Форма) Экспорт

	ИмяТекущегоЭлемента = "СтрокаПоискаНоменклатура";
	Возврат ИмяТекущегоЭлемента;

КонецФункции

// Процедура устанавливает текущим элементом строку поиска на форме.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма подбора.
//
Процедура УстановитьТекущийЭлементСтрокаПоиска(Форма) Экспорт
	
	ВариантПоиска = ПодборТоваровКлиентСервер.ДействующийВариантПоиска(Форма);
	Если Не ВариантПоиска = "Расширенный" Тогда
		Возврат;
	КонецЕсли;
	
	ИмяТекущегоЭлемента = ИмяТекущегоЭлементаСтрокиПоиска(Форма);
	
	Форма.ТекущийЭлемент = Форма.Элементы[ИмяТекущегоЭлемента];
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовФормы

// Процедура вызывается при активизации строки списка иерархии номенклатуры.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма списка, форма подборов.
//
Процедура ПриАктивизацииСтрокиИерархииНоменклатуры(Форма) Экспорт
	
	Если Не Форма.ИспользоватьФильтры Тогда
		Возврат;
	КонецЕсли;
	
	Если Не (Форма.ВариантНавигации = "ПоИерархии") Тогда
		Возврат;
	КонецЕсли;
	
	Форма.ПодключитьОбработчикОжидания("ИерархияНоменклатурыПриАктивизацииСтрокиОбработчикОжидания", 0.1, Истина);
	
КонецПроцедуры

// Процедура вызывается при активизации строки списка иерархии номенклатуры.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма списка, форма подборов.
//
Процедура ОбработчикАктивизацииСтрокиИерархииНоменклатуры(Форма) Экспорт

	Если Форма.ТекущаяИерархияНоменклатуры = Форма.Элементы.ИерархияНоменклатуры.ТекущаяСтрока Тогда
		Возврат;
	КонецЕсли;
	
	Форма.ТекущаяИерархияНоменклатуры = Форма.Элементы.ИерархияНоменклатуры.ТекущаяСтрока;
	
	УстановитьОтборПоИерархииНоменклатуры(Форма);

КонецПроцедуры

// Устанавливает отбор по группе номенклатыр в списке номенклатуры.
// Если на форме используется вариант навигации отличный от навигации по иерархии,
// то отбор не устанавливается.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма списка номенклатуры или форма подбора.
//
Процедура УстановитьОтборПоИерархииНоменклатуры(Форма)
	
	Если Не Форма.ИспользоватьФильтры Тогда
		Возврат;
	КонецЕсли;
	
	Если Не Форма.ВариантНавигации = "ПоИерархии" Тогда
		Возврат;
	КонецЕсли;
	
	Если Форма.Элементы.ИерархияНоменклатуры.ТекущаяСтрока = Неопределено Тогда
		ОтборПоИерархииНоменклатуры = ПредопределенноеЗначение("Справочник.Номенклатура.ПустаяСсылка");
	Иначе
		ОтборПоИерархииНоменклатуры = Форма.Элементы.ИерархияНоменклатуры.ТекущаяСтрока;
	КонецЕсли;
	
	ОбщегоНазначенияКлиентСервер.УстановитьЭлементОтбораДинамическогоСписка(Форма.СписокНоменклатура, "Родитель", ОтборПоИерархииНоменклатуры, 
		ВидСравненияКомпоновкиДанных.Равно, "Родитель", Истина);
	
КонецПроцедуры

// Процедура вызывается при активизации строки списка номенклатуры.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма списка.
//
Процедура ПриАктивизацииСтрокиСпискаНоменклатуры(Форма) Экспорт
	
	Форма.ПодключитьОбработчикОжидания("СписокПриАктивизацииСтрокиОбработчикОжидания", 0.1, Истина);
	
КонецПроцедуры

#КонецОбласти

#Область ПоискНаФормахПодбора

// Если используется полнотекстовый поиск, то функция проверяет актуальность индекса.
// Если индекс не актуален, то отображается диалог с предложением обновить индекс
// полнотекстового поиска. Если пользователь отказывается от обновления индекса,
// то все равно разрешается выполнить расширенный поиск, т.к. допускается что сведения
// касающие товаров часто не меняются.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма из которой выполняется вызов функции поиска.
//  ОписаниеОповещенияПриУспехе - ОписаниеОповещения - описание оповещения.
//
// Возвращаемое значение:
//	 Булево - Истина - выполнение поиска товаров возможно, Ложь - нет.
//
Процедура ВыполнениеРасширенногоПоискаВозможно(Форма, ОписаниеОповещенияПриУспехе = Неопределено) Экспорт
	
	// При включении флажка "Поиск по точному соответствию", выполняется неполнотекстовый поиск, 
	// соответственно, не нужно проверять актуальность индекса полнотекстового поиска.
	Если Форма.НайтиНоменклатуруПоТочномуСоответствию Тогда
		Если ОписаниеОповещенияПриУспехе <> Неопределено Тогда
			ВыполнитьОбработкуОповещения(ОписаниеОповещенияПриУспехе);
		КонецЕсли;
		Возврат;
	КонецЕсли;
	
	// Проверка необходима только при использовании полнотекстового поиска.
	Если Не Форма.ИспользоватьПолнотекстовыйПоиск Тогда
		Если ОписаниеОповещенияПриУспехе <> Неопределено Тогда
			ВыполнитьОбработкуОповещения(ОписаниеОповещенияПриУспехе);
		КонецЕсли;
		Возврат;
	КонецЕсли;
	
	Если Форма.ИнформационнаяБазаФайловая И Не Форма.ИндексПолнотекстовогоПоискаАктуален Тогда
		ДополнительныеПараметры = Новый Структура("Форма, ОписаниеОповещенияПриУспехе", Форма, ОписаниеОповещенияПриУспехе);
		Оповещение = Новый ОписаниеОповещения("ВыполнениеРасширенногоПоискаВозможноПослеВопросаОбИндексах", ЭтотОбъект, ДополнительныеПараметры);
		ПоказатьВопрос(Оповещение, НСтр("ru = 'Индекс полнотекстового поиска неактуален. Обновить индекс?'"), РежимДиалогаВопрос.ДаНет);
	Иначе
		Если ОписаниеОповещенияПриУспехе <> Неопределено Тогда
			ВыполнитьОбработкуОповещения(ОписаниеОповещенияПриУспехе);
		КонецЕсли;
	КонецЕсли;
		
КонецПроцедуры

// Продолжение предыдущей процедуры.
// Параметры:
//		Результат - КодВозвратаДиалога - ответ на вопрос.
//		ДополнительныеПараметры - Произвольный - доп. параметры.
Процедура ВыполнениеРасширенногоПоискаВозможноПослеВопросаОбИндексах(Результат, ДополнительныеПараметры) Экспорт
	
	Форма = ДополнительныеПараметры.Форма;
	ОписаниеОповещенияПриУспехе = ДополнительныеПараметры.ОписаниеОповещенияПриУспехе;
	
	Если Результат = КодВозвратаДиалога.Да Тогда
		ОбновитьИндексПолнотекстовогоПоиска();
		Форма.ИндексПолнотекстовогоПоискаАктуален = Истина;
	КонецЕсли;
	
	Если ОписаниеОповещенияПриУспехе <> Неопределено Тогда
		ВыполнитьОбработкуОповещения(ОписаниеОповещенияПриУспехе);
	КонецЕсли;
	
КонецПроцедуры

// Обрабатывает флаг возврата функции выполнения поиска номенклатуры.
// Если поиск не был выполнен, то выводится предупреждение.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма из которой вызывалась функция поиска товаров по строке,
//	 ОтображатьПредупреждение - Булево - флаг отображения предупреждения, при неудачном поиске.
//
Процедура ПослеВыполненияПоискаНоменклатуры(Форма, ОтображатьПредупреждение = Истина) Экспорт
	
	ПоискВыполнен = Не Форма.ПоискНоменклатурыНеУдачный;
	КодОшибки = Форма.КодОшибкиПоиска;
	
	СтрокаПоиска = Форма.СтрокаПоискаНоменклатура;
	
	Если ЗначениеЗаполнено(СтрокаПоиска) Тогда
		Если Не ПоискВыполнен И ОтображатьПредупреждение Тогда
			ПоказатьПредупреждение(Неопределено,ТекстПредупрежденияОшибкиРасширенногоПоиска(КодОшибки), 120, "Поиск");
		КонецЕсли;
        ОбновитьСписокВыбора(Форма.Элементы.СтрокаПоискаНоменклатура.СписокВыбора, СтрокаПоиска, 21);      
    КонецЕсли;
	
	// Установить текущий элемент формы.
	Если Не ПодборТоваровКлиентСервер.ЭтоФормаПодбораТоваровПоКатегориям(Форма) Тогда
		ТекущийЭлемент = Форма.Элементы[ПодборТоваровКлиентСервер.ИмяСпискаНоменклатурыПоВариантуПоиска(Форма)];
	Иначе
		ТекущийЭлемент = Форма.Элементы["СписокРасширенныйПоискНоменклатура"];
	КонецЕсли;
	
	Если Не ПоискВыполнен Тогда
		ТекущийЭлемент = Форма.Элементы.СтрокаПоискаНоменклатура;
	КонецЕсли;
	
	Форма.ТекущийЭлемент = ТекущийЭлемент;
	
КонецПроцедуры

// Обновляет список выбора.
// Параметры:
//		СписокВыбора - СписокЗначений - список выбора.
//		СтрокаПоиска - Строка - строка поиска.
//		РазмерСпискаВыбора - Число - количество элементов в списке выбора.
Процедура ОбновитьСписокВыбора(СписокВыбора, СтрокаПоиска, РазмерСпискаВыбора = 1000) Экспорт
	
	// Удалим элемент из истории поиска если он там был
	НомерНайденногоЭлементаСписка = СписокВыбора.НайтиПоЗначению(СтрокаПоиска);
	Пока НомерНайденногоЭлементаСписка <> Неопределено Цикл
		СписокВыбора.Удалить(НомерНайденногоЭлементаСписка);
		НомерНайденногоЭлементаСписка = СписокВыбора.НайтиПоЗначению(СтрокаПоиска);
	КонецЦикла;
	
	
	// И поставим его на первое место
	СписокВыбора.Вставить(0, СтрокаПоиска);
	Пока СписокВыбора.Количество() > РазмерСпискаВыбора Цикл
		СписокВыбора.Удалить(СписокВыбора.Количество() - 1);
	КонецЦикла;
	
КонецПроцедуры // ОбновитьСписокВыбора()

// Возвращает текст предупреждения по коду ошибки расширенного поиска.
//
// Параметры:
//	 КодОшибки - Строка - код ошибки расширенного поиска.
//
// Возвращаемое значение:
//	 Строка - текст предупреждения ошибки расширеннго поиска.
//
Функция ТекстПредупрежденияОшибкиРасширенногоПоиска(КодОшибки)
	
	Если КодОшибки = "НичегоНеНайдено" Тогда
		ТекстПредупреждения = НСтр("ru = 'Ничего не найдено, уточните запрос.'");
	ИначеЕсли КодОшибки = "СлишкомМногоРезультатов" Тогда
		ТекстПредупреждения = НСтр("ru = 'Слишком много результатов поиска, уточните запрос.'");
	ИначеЕсли КодОшибки = "ОшибкаПоиска" Тогда
		ТекстПредупреждения = НСтр("ru = 'При выполнении поиска произошла ошибка, попробуйте изменить выражение поиска.'");
	Иначе
		ТекстПредупреждения = "";
	КонецЕсли;
	
	Возврат ТекстПредупреждения;
	
КонецФункции

// Вызывает функцию обновления индекса полнотекстового поиска на сервере.
//
Процедура ОбновитьИндексПолнотекстовогоПоиска() Экспорт
	
	Состояние(НСтр("ru = 'Идет обновление индекса полнотекстового поиска ...'"));
	
	ПодборТоваровВызовСервера.ОбновитьИндексПолнотекстовогоПоиска();
	
	Состояние(НСтр("ru = 'Обновление индекса полнотекстового поиска завершено.'"));
	
КонецПроцедуры

#КонецОбласти

#Область ПроцедурыНавигации

// Устанавливает текущую строку иерархии номенклатуры в формах списков номенклатуры и подборов.
//
// Параметры:
//	 Форма - ФормаКлиентскогоПриложения - форма списка, форма подбора.
//
Процедура УстановитьТекущуюСтрокуИерархииНоменклатуры(Форма) Экспорт
	
	Если Форма.ИспользоватьФильтры Тогда
		Возврат;
	КонецЕсли;
	
	ЭтоФормаПодбора = ПодборТоваровКлиентСервер.ЭтоФормаПодбора(Форма);
	
	ИмяСпискаНоменклатуры = ПодборТоваровКлиентСервер.ИмяСпискаНоменклатурыПоВариантуПоиска(Форма);
	
	ТекущаяСтрока = Форма.Элементы[ИмяСпискаНоменклатуры].ТекущаяСтрока;
	Если ТекущаяСтрока = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Если ТипЗнч(ТекущаяСтрока) = Тип("СтрокаГруппировкиДинамическогоСписка") Тогда
		Возврат;
	КонецЕсли;
	
	Если ЭтоФормаПодбора И Форма.НавигацияПоХарактеристикам Тогда
		Возврат;
	КонецЕсли;
	
	ТекущиеДанные = Форма.Элементы[ИмяСпискаНоменклатуры].ТекущиеДанные;
	
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Если Форма.Элементы.ИерархияНоменклатуры.ТекущаяСтрока = ТекущиеДанные.Родитель Тогда
		Возврат;
	КонецЕсли;
	
	Форма.Элементы.ИерархияНоменклатуры.ТекущаяСтрока = ТекущиеДанные.Родитель;
	
КонецПроцедуры

#КонецОбласти

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ПриАктивизацииСтрокиСпискаВидыНоменклатуры(Форма) Экспорт
	
	Если Форма.ВариантНавигации = "ПоИерархии" Тогда
		Возврат;
	КонецЕсли;
	
	ТекущиеДанные = Форма.Элементы.ВидыНоменклатуры.ТекущиеДанные;
	
	Если ТекущиеДанные = Неопределено Или Форма.ВидНоменклатуры = ТекущиеДанные.Ссылка Тогда
		Возврат;
	КонецЕсли;
	
	Форма.ВидНоменклатуры = ТекущиеДанные.Ссылка;

	Если Не Форма.ИспользоватьФильтры Тогда
		Возврат;
	КонецЕсли;
		
	Форма.ПодключитьОбработчикОжидания("ВидыНоменклатурыПриАктивизацииСтрокиОбработчикОжидания", 0.1, Истина);
КонецПроцедуры

#КонецОбласти
