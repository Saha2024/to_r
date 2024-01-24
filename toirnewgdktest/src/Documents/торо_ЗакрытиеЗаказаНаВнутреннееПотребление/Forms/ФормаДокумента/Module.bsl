&НаКлиенте
Перем КэшированныеЗначения; // используется механизмом обработки изменения реквизитов ТЧ

#Область ОбработчикиСобытийФормы 

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	Если Параметры.Свойство("АвтоТест") Тогда
		Возврат;
	КонецЕсли;
	
	// СтандартныеПодсистемы.Свойства
	ДополнительныеПараметры = Новый Структура;
	ДополнительныеПараметры.Вставить("ИмяЭлементаДляРазмещения", "ГруппаДополнительныеРеквизиты");
	УправлениеСвойствами.ПриСозданииНаСервере(ЭтотОбъект, ДополнительныеПараметры);
	// Конец СтандартныеПодсистемы.Свойства

	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКоманды.ПриСозданииНаСервере(ЭтаФорма);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	// СтандартныеПодсистемы.ВерсионированиеОбъектов
	ВерсионированиеОбъектов.ПриСозданииНаСервере(ЭтаФорма);
	// Конец СтандартныеПодсистемы.ВерсионированиеОбъектов

	Если Не ЗначениеЗаполнено(Объект.Ссылка) Тогда
		ПриЧтенииСозданииНаСервере();
	КонецЕсли;
		
	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
	
	УстановитьУсловноеОформление();
	
	Если Параметры.ЗначенияЗаполнения.Свойство("ИзРабочегоМестаМТО") Тогда
		ИзРабочегоМестаМТО = Параметры.ЗначенияЗаполнения.ИзРабочегоМестаМТО;
		Если ИзРабочегоМестаМТО Тогда 
			ПолучитьТабличнуюЧастьПоЗаказу(Объект.ЗаказНаВнутреннееПотребление);
		КонецЕсли;
	Иначе
		ИзРабочегоМестаМТО = Ложь;
	КонецЕсли;
КонецПроцедуры

&НаСервере
Процедура ПриЧтенииНаСервере(ТекущийОбъект)
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиентСервер.ОбновитьКоманды(ЭтотОбъект, Объект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	// СтандартныеПодсистемы.УправлениеДоступом
	УправлениеДоступом.ПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.УправлениеДоступом
	
	ПриЧтенииСозданииНаСервере();
КонецПроцедуры

&НаСервере
Процедура ПриЧтенииСозданииНаСервере()
	ЗаполнитьСлужебныеРеквизитыПоНоменклатуре();
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	// СтандартныеПодсистемы.Свойства
    УправлениеСвойствамиКлиент.ПослеЗагрузкиДополнительныхРеквизитов(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиент.НачатьОбновлениеКоманд(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	Если ИзРабочегоМестаМТО И НЕ ЗначениеЗаполнено(Объект.Ссылка) Тогда 
		Модифицированность = Истина;
	КонецЕсли;
КонецПроцедуры

&НаСервере
Процедура ОбработкаПроверкиЗаполненияНаСервере(Отказ, ПроверяемыеРеквизиты)
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ОбработкаПроверкиЗаполнения(ЭтотОбъект, Отказ, ПроверяемыеРеквизиты);
	// Конец СтандартныеПодсистемы.Свойства
КонецПроцедуры

&НаСервере
Процедура ПередЗаписьюНаСервере(Отказ, ТекущийОбъект, ПараметрыЗаписи)
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ПередЗаписьюНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.Свойства
КонецПроцедуры

&НаСервере
Процедура ПослеЗаписиНаСервере(ТекущийОбъект, ПараметрыЗаписи)
	// СтандартныеПодсистемы.УправлениеДоступом
	УправлениеДоступом.ПослеЗаписиНаСервере(ЭтотОбъект, ТекущийОбъект, ПараметрыЗаписи);
	// Конец СтандартныеПодсистемы.УправлениеДоступом
	
	ЗаполнитьСлужебныеРеквизитыПоНоменклатуре();
	
	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
КонецПроцедуры

&НаКлиенте
Процедура ОбработкаОповещения(ИмяСобытия, Параметр, Источник)
	// СтандартныеПодсистемы.Свойства 
	Если УправлениеСвойствамиКлиент.ОбрабатыватьОповещения(ЭтотОбъект, ИмяСобытия, Параметр) Тогда
		ОбновитьЭлементыДополнительныхРеквизитов();
		УправлениеСвойствамиКлиент.ПослеЗагрузкиДополнительныхРеквизитов(ЭтотОбъект);
	КонецЕсли;
	// Конец СтандартныеПодсистемы.Свойства
КонецПроцедуры

&НаКлиенте
Процедура ПослеЗаписи(ПараметрыЗаписи)
	
	Если ИзРабочегоМестаМТО Тогда 
		ПараметрыОбновления = Новый Структура("Проведен", ПараметрыЗаписи.РежимЗаписи <> РежимЗаписиДокумента.Запись);
		Оповестить("ОбновитьМТО", ПараметрыОбновления); 
	КонецЕсли;
	
	ПодключаемыеКомандыКлиент.ПослеЗаписи(ЭтотОбъект, Объект, ПараметрыЗаписи);

КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ЗаказОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	Если ВыбранноеЗначение <> Объект.ЗаказНаВнутреннееПотребление Тогда
		Если ЗначениеЗаполнено(Объект.Товары) Тогда
			СтандартнаяОбработка = Ложь; 
			ОповещениеОбОтвете = Новый ОписаниеОповещения("ЗавершитьВыборЗаказа", ЭтаФорма, Новый Структура("НовыйЗаказ", ВыбранноеЗначение));
			ПоказатьВопрос(ОповещениеОбОтвете, 
				НСтр("ru = 'Табличная часть будет перезаполнена по выбранному заказу на внутреннее потребление. Продолжить?'"), 
				РежимДиалогаВопрос.ДаНет);
		Иначе			
			ПолучитьТабличнуюЧастьПоЗаказу(ВыбранноеЗначение);
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура КомментарийНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
    СтандартнаяОбработка = Ложь;
	ОбщегоНазначенияКлиент.ПоказатьФормуРедактированияКомментария(
		Элемент.ТекстРедактирования, ЭтотОбъект, "Объект.Комментарий");		
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыТовары

&НаКлиенте
Процедура ТоварыНоменклатураПриИзменении(Элемент)
	ТекущаяСтрока = Элементы.Товары.ТекущиеДанные;
	
	СтруктураДействий = Новый Структура;
	СтруктураДействий.Вставить("ПроверитьХарактеристикуПоВладельцу", ТекущаяСтрока.Характеристика);
	СтруктураДействий.Вставить("НоменклатураПриИзмененииПереопределяемый", Новый Структура("ИмяФормы, ИмяТабличнойЧасти",
		ЭтаФорма.ИмяФормы, "Товары"));

	ОбработкаТабличнойЧастиКлиент.ОбработатьСтрокуТЧ(ТекущаяСтрока, СтруктураДействий, КэшированныеЗначения);
	ТекущаяСтрока.ХарактеристикиИспользуются =  торо_НоменклатураСервер.ПолучитьХарактеристикиИспользуются(ТекущаяСтрока.Номенклатура);
	ТекущаяСтрока.СерииИспользуются = ПолучитьСерииИспользуются(ТекущаяСтрока.Номенклатура);
	ТекущаяСтрока.Серия = Неопределено;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыХарактеристикаСоздание(Элемент, СтандартнаяОбработка)
	Если Элементы.Товары.ТекущиеДанные <> Неопределено И ЗначениеЗаполнено(Элементы.Товары.ТекущиеДанные.Номенклатура) Тогда
		Вид = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(Элементы.Товары.ТекущиеДанные.Номенклатура, "ВидНоменклатуры");
		СтруктураПараметров = Новый Структура("ВидНоменклатуры, Владелец", Вид, Элементы.Товары.ТекущиеДанные.Номенклатура);
		ОткрытьФорму("Справочник.ХарактеристикиНоменклатуры.Форма.ФормаЭлемента", СтруктураПараметров);
		СтандартнаяОбработка = Ложь;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыСерияСоздание(Элемент, СтандартнаяОбработка)
	Если Элементы.Товары.ТекущиеДанные <> Неопределено И ЗначениеЗаполнено(Элементы.Товары.ТекущиеДанные.Номенклатура) Тогда
		СтруктураПараметров = Новый Структура("ВидНоменклатуры", торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(Элементы.Товары.ТекущиеДанные.Номенклатура, "ВидНоменклатуры"));
		ОткрытьФорму("Справочник.СерииНоменклатуры.Форма.ФормаЭлемента", СтруктураПараметров);
		СтандартнаяОбработка = Ложь;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыХарактеристикаОкончаниеВводаТекста(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, СтандартнаяОбработка)
	ТекущиеДанные = Элементы.Товары.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено Тогда
		ВидНоменклатуры = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Номенклатура, "ВидНоменклатуры");
		ИспользованиеХарактеристик = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ВидНоменклатуры, "ИспользованиеХарактеристик");
		ПараметрыПолученияДанных.Отбор.Очистить();
		Если ИспользованиеХарактеристик = ПредопределенноеЗначение("Перечисление.ВариантыИспользованияХарактеристикНоменклатуры.ОбщиеДляВидаНоменклатуры") Тогда
			ПараметрыПолученияДанных.Отбор.Вставить("Владелец", ВидНоменклатуры);
		Иначе
			ПараметрыПолученияДанных.Отбор.Вставить("Владелец", ТекущиеДанные.Номенклатура);
		КонецЕсли; 
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыХарактеристикаАвтоПодбор(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, Ожидание, СтандартнаяОбработка)
	ТекущиеДанные = Элементы.Товары.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено Тогда
		ВидНоменклатуры = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Номенклатура, "ВидНоменклатуры");
		ИспользованиеХарактеристик = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ВидНоменклатуры, "ИспользованиеХарактеристик");
		ПараметрыПолученияДанных.Отбор.Очистить();
		Если ИспользованиеХарактеристик = ПредопределенноеЗначение("Перечисление.ВариантыИспользованияХарактеристикНоменклатуры.ОбщиеДляВидаНоменклатуры") Тогда
			ПараметрыПолученияДанных.Отбор.Вставить("Владелец", ВидНоменклатуры);
		Иначе
			ПараметрыПолученияДанных.Отбор.Вставить("Владелец", ТекущиеДанные.Номенклатура);
		КонецЕсли; 
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыСерияАвтоПодбор(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, Ожидание, СтандартнаяОбработка)
	Если Ожидание > 0 Тогда
		ТекущиеДанные = Элементы.Товары.ТекущиеДанные;
		ВидНоменклатуры = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Номенклатура, "ВидНоменклатуры");
		ПараметрыПолученияДанных.Отбор.Вставить("ВидНоменклатуры",ВидНоменклатуры);
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыСерияОкончаниеВводаТекста(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, СтандартнаяОбработка)
	ТекущиеДанные = Элементы.Товары.ТекущиеДанные;
	ВидНоменклатуры = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Номенклатура, "ВидНоменклатуры");
	ПараметрыПолученияДанных.Отбор.Вставить("ВидНоменклатуры", ВидНоменклатуры);
КонецПроцедуры

&НаКлиенте
Процедура ТоварыСерияОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	ТекущиеДанные = Элементы.Товары.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено И ЗначениеЗаполнено(ВыбранноеЗначение) Тогда
		ВидНоменклатуры = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Номенклатура, "ВидНоменклатуры");
		ВидНоменклатурыВыбранный = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ВыбранноеЗначение, "ВидНоменклатуры");
		Если ВидНоменклатуры <> ВидНоменклатурыВыбранный Тогда
			СтандартнаяОбработка = Ложь;
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ТоварыСерияНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	ТекДанные =  Элементы.Товары.ТекущиеДанные;
	Если Не ТекДанные = Неопределено Тогда
		СтандартнаяОбработка = Ложь;
		ВидНоменклатуры = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекДанные.Номенклатура, "ВидНоменклатуры");
		ПараметрыФормы = Новый Структура("Отбор", Новый Структура("ВидНоменклатуры", ВидНоменклатуры));
		ОткрытьФорму("Справочник.СерииНоменклатуры.ФормаВыбора", ПараметрыФормы, Элемент);
	КонецЕсли;
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

// СтандартныеПодсистемы.Свойства
&НаКлиенте
Процедура Подключаемый_СвойстваВыполнитьКоманду(ЭлементИлиКоманда, НавигационнаяСсылка = Неопределено, СтандартнаяОбработка = Неопределено)
	УправлениеСвойствамиКлиент.ВыполнитьКоманду(ЭтотОбъект, ЭлементИлиКоманда, СтандартнаяОбработка);
КонецПроцедуры
// Конец СтандартныеПодсистемы.Свойства

// СтандартныеПодсистемы.ПодключаемыеКоманды

&НаКлиенте
Процедура Подключаемый_ВыполнитьКоманду(Команда)
     ПодключаемыеКомандыКлиент.НачатьВыполнениеКоманды(ЭтотОбъект, Команда, Объект);
КонецПроцедуры
 
&НаКлиенте
Процедура Подключаемый_ПродолжитьВыполнениеКомандыНаСервере(ПараметрыВыполнения, ДополнительныеПараметры) Экспорт
     ВыполнитьКомандуНаСервере(ПараметрыВыполнения);
КонецПроцедуры
 
&НаСервере
Процедура ВыполнитьКомандуНаСервере(ПараметрыВыполнения)
     ПодключаемыеКоманды.ВыполнитьКоманду(ЭтотОбъект, ПараметрыВыполнения, Объект);
КонецПроцедуры
 
&НаКлиенте
Процедура Подключаемый_ОбновитьКоманды()
     ПодключаемыеКомандыКлиентСервер.ОбновитьКоманды(ЭтотОбъект, Объект);
КонецПроцедуры

// Конец СтандартныеПодсистемы.ПодключаемыеКоманды

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура ЗаполнитьСлужебныеРеквизитыПоНоменклатуре()
	
	НоменклатураСервер.ЗаполнитьСлужебныеРеквизитыПоНоменклатуреВКоллекции(
		Объект.Товары,
		Новый Структура("ЗаполнитьПризнакХарактеристикиИспользуются, ЗаполнитьПризнакТипНоменклатуры, ЗаполнитьПризнакАртикул, ЗаполнитьПризнакСерииИспользуются",
			Новый Структура("Номенклатура", "ХарактеристикиИспользуются"),
			Новый Структура("Номенклатура", "ТипНоменклатуры"),
			Новый Структура("Номенклатура", "Артикул"),
			Новый Структура("Номенклатура", "СерииИспользуются")));
	
КонецПроцедуры

// СтандартныеПодсистемы.Свойства 
&НаСервере
Процедура ОбновитьЭлементыДополнительныхРеквизитов()
	УправлениеСвойствами.ОбновитьЭлементыДополнительныхРеквизитов(ЭтотОбъект);
КонецПроцедуры

&НаКлиенте
Процедура ОбновитьЗависимостиДополнительныхРеквизитов()
	УправлениеСвойствамиКлиент.ОбновитьЗависимостиДополнительныхРеквизитов(ЭтотОбъект);
КонецПроцедуры

&НаКлиенте
Процедура Подключаемый_ПриИзмененииДополнительногоРеквизита(Элемент)
	УправлениеСвойствамиКлиент.ОбновитьЗависимостиДополнительныхРеквизитов(ЭтотОбъект);
КонецПроцедуры
// Конец СтандартныеПодсистемы.Свойства

&НаСервере
Процедура УстановитьУсловноеОформление() 
	// Оформление серий
	Элемент = УсловноеОформление.Элементы.Добавить();
	
	ПолеЭлемента = Элемент.Поля.Элементы.Добавить();
	ПолеЭлемента.Поле = Новый ПолеКомпоновкиДанных(Элементы.ТоварыСерия.Имя);
	
	ОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("Объект.Товары.СерииИспользуются");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
	ОтборЭлемента.ПравоеЗначение = Ложь;
	
	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста", WebЦвета.НейтральноСерый);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст", НСтр("ru = '<серии не используются>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ОтметкаНезаполненного", Ложь);
	Элемент.Оформление.УстановитьЗначениеПараметра("Доступность", Ложь);
		
	// Оформление характеристик
	Элемент = УсловноеОформление.Элементы.Добавить();

	ПолеЭлемента = Элемент.Поля.Элементы.Добавить();
	ПолеЭлемента.Поле = Новый ПолеКомпоновкиДанных(Элементы.ТоварыХарактеристика.Имя);

	ОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("Объект.Товары.ХарактеристикиИспользуются");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
	ОтборЭлемента.ПравоеЗначение = Ложь;

	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста", ЦветаСтиля.ТекстЗапрещеннойЯчейкиЦвет);
	Элемент.Оформление.УстановитьЗначениеПараметра("ОтметкаНезаполненного", Ложь);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст", НСтр("ru = '<характеристики не используются>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ТолькоПросмотр", Истина);
КонецПроцедуры

&НаСервере
Функция ПолучитьСерииИспользуются(Номенклатура)
	Возврат Номенклатура.ВидНоменклатуры.ИспользоватьСерии;
КонецФункции

&НаКлиенте
Процедура ЗавершитьВыборЗаказа(Результат, ДопПараметры) Экспорт
	Если Результат = КодВозвратаДиалога.Да Тогда
		Ошибка = Ложь;
		ПолучитьТабличнуюЧастьПоЗаказу(ДопПараметры.НовыйЗаказ, Ошибка);
		Объект.ЗаказНаВнутреннееПотребление = ДопПараметры.НовыйЗаказ;
		Если Ошибка Тогда 
			Объект.Товары.Очистить();
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры

&НаСервере
Процедура ПолучитьТабличнуюЧастьПоЗаказу(ЗаказНаВнутреннееПотребление, Ошибка = Ложь)
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	Заказы.Номенклатура КАК Номенклатура,
	               |	Заказы.Характеристика КАК Характеристика,
	               |	Заказы.Серия КАК Серия
	               |ИЗ
	               |	РегистрНакопления.ЗаказыНаВнутреннееПотребление.Остатки(, ЗаказНаВнутреннееПотребление = &ЗаказНаВнутреннееПотребление) КАК Заказы";
	Запрос.УстановитьПараметр("ЗаказНаВнутреннееПотребление", ЗаказНаВнутреннееПотребление);
	Если НЕ Константы.торо_ИспользоватьХарактеристикиНоменклатуры.Получить() Тогда 
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "Заказы.Характеристика", """""");
	КонецЕсли;
	
	Если НЕ Константы.ИспользоватьСерииНоменклатуры.Получить() Тогда
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "Заказы.Серия", """""");
	КонецЕсли;
	
	РезЗапроса = Запрос.Выполнить();
	Если РезЗапроса.Пустой() Тогда
		ОбщегоНазначения.СообщитьПользователю(НСтр("ru = 'Для выбранного заказа на внутреннее потребление потребность закрыта.'"));
		Ошибка = Истина;
	Иначе
		Объект.Товары.Загрузить(РезЗапроса.Выгрузить());
		ЗаполнитьСлужебныеРеквизитыПоНоменклатуре();
	КонецЕсли;
КонецПроцедуры

#КонецОбласти