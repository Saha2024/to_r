#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
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
	
	торо_СобытияФорм.ПриСозданииНаСервере(ЭтаФорма, Отказ, СтандартнаяОбработка, Новый Структура("УстановитьСвойствоЭлементовФормыОтПрав", Истина));
	
	ТекСтруктураИерархии = 	ТекСтруктураИерархии = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить(
			"НастройкиТОиР",
			"ОсновнаяСтруктураИерархии",
			Истина);

	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
	
	торо_УправлениеДоступом.УстановитьДоступностьРедактированияПоОграничениюДоступаДляОР(ЭтаФорма, Объект.ОбъектыРемонта.Выгрузить());
	
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

КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	// СтандартныеПодсистемы.Свойства
    УправлениеСвойствамиКлиент.ПослеЗагрузкиДополнительныхРеквизитов(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиент.НачатьОбновлениеКоманд(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
КонецПроцедуры

&НаСервере
Процедура ОбработкаПроверкиЗаполненияНаСервере(Отказ, ПроверяемыеРеквизиты)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ОбработкаПроверкиЗаполнения(ЭтотОбъект, Отказ, ПроверяемыеРеквизиты);
	// Конец СтандартныеПодсистемы.Свойства
	
	Для Каждого Строка Из Объект.ОбъектыРемонта Цикл  
		Если НЕ ЗначениеЗаполнено(Строка.ОбъектРемонта) Тогда
			ТекстСообщения = СтрШаблон(НСтр("ru = 'В строке %1 не заполнен объект ремонта.'"), Строка.НомерСтроки);
			Поле = СтрШаблон("Объект.ОбъектыРемонта[%1].ОбъектРемонта", Строка.НомерСтроки - 1);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
		КонецЕсли;
	КонецЦикла;
	
	Если Отказ Тогда 
		Возврат;
	КонецЕсли;
	
	ШаблонСообщения = Нстр("ru='Для объекта ремонта ""%1"" не заполнено поле ""%2"".'");
	Для каждого СтрокаКоэффициента из Объект.КоэффициентыРемонтныхОсобенностей Цикл
		Если НЕ ЗначениеЗаполнено(СтрокаКоэффициента.ДатаНачалаИспользования) Тогда
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаКоэффициента.ОбъектРемонта, НСтр("ru='дата ввода'"));
			Поле = СтрШаблон("Объект.КоэффициентыРемонтныхОсобенностей[%1].ДатаНачалаИспользования", Объект.КоэффициентыРемонтныхОсобенностей.Индекс(СтрокаКоэффициента));
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
		КонецЕсли;
		Если НЕ ЗначениеЗаполнено(СтрокаКоэффициента.ВидКоэффициента) Тогда
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаКоэффициента.ОбъектРемонта, НСтр("ru='вид коэффициента'"));
			Поле = СтрШаблон("Объект.КоэффициентыРемонтныхОсобенностей[%1].ВидКоэффициента", Объект.КоэффициентыРемонтныхОсобенностей.Индекс(СтрокаКоэффициента));
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,, Поле,, Отказ);
		КонецЕсли;
	КонецЦикла;

	
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
	
	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
	
КонецПроцедуры

&НаКлиенте
Процедура ПослеЗаписи(ПараметрыЗаписи)
	
	ПодключаемыеКомандыКлиент.ПослеЗаписи(ЭтотОбъект, Объект, ПараметрыЗаписи);
	
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

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыОбъектыРемонта
&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаПриИзменении(Элемент)
	
	Если НЕ ДобавлениеОР Тогда
		МассивСтрок = Объект.КоэффициентыРемонтныхОсобенностей.НайтиСтроки(Новый Структура("ОбъектРемонта", СтарыйОР));
		Если НЕ Элементы.ОбъектыРемонта.ТекущиеДанные = Неопределено Тогда
			НовыйОР = Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта;
		КонецЕсли;
		
		Для каждого Элем Из МассивСтрок Цикл
			Элем.ОбъектРемонта = НовыйОР;
		КонецЦикла;
	КонецЕсли;
	
	ОбъектыРемонтаПриАктивизацииСтроки(Элементы.ОбъектыРемонта);
	Элементы.ОбъектыРемонта.ЗакончитьРедактированиеСтроки(Ложь);
	ДобавлениеОР = Ложь;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаПриАктивизацииСтроки(Элемент)
	
	Если НЕ Элементы.ОбъектыРемонта.ТекущиеДанные = Неопределено Тогда
		Если ЗначениеЗаполнено(Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта) Тогда
			Элементы.КоэффициентыРемонтныхОсобенностей.ОтборСтрок = Новый ФиксированнаяСтруктура("ОбъектРемонта", Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта);
			СтарыйОР = Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта;
		Иначе
			Элементы.КоэффициентыРемонтныхОсобенностей.ОтборСтрок = Новый ФиксированнаяСтруктура("ОбъектРемонта", ПредопределенноеЗначение("Справочник.торо_ОбъектыРемонта.ПустаяСсылка"));
		КонецЕсли;
			
		Если ТипЗнч(Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта) = ТИп("СправочникСсылка.торо_СписокОбъектовРегламентногоМероприятия") Тогда
			Элементы.КоэффициентыРемонтныхОсобенностейИерархия.Видимость = Ложь;
			Элементы.КоэффициентыРемонтныхОсобенностейРаспространятьНаПодчиненных.Видимость = Ложь;
		Иначе
			Элементы.КоэффициентыРемонтныхОсобенностейИерархия.Видимость = Истина;
			Элементы.КоэффициентыРемонтныхОсобенностейРаспространятьНаПодчиненных.Видимость = Истина;
		КонецЕсли;
		
	Иначе
		Элементы.КоэффициентыРемонтныхОсобенностей.ОтборСтрок = Новый ФиксированнаяСтруктура("ОбъектРемонта", Неопределено);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	ТекДанные = Элементы.ОбъектыРемонта.ТекущиеДанные;
	
	ОбъектыРемонтаПриАктивизацииСтроки(Элементы.ОбъектыРемонта);
	
	Если НЕ ТекДанные = Неопределено Тогда
		Если ТипЗнч(ТекДанные.ОбъектРемонта) = Тип("СправочникСсылка.торо_ОбъектыРемонта") Тогда
			СтандартнаяОбработка = Ложь;     
			СписокДоступныхСтатусов = торо_СтатусыОРВУчете.СписокСтатусовДляПодбора();
			СтарыйОР = ТекДанные.ОбъектРемонта;
			
			ПараметрыОтбора = Новый Структура("СписокСтатусов", СписокДоступныхСтатусов);
			ПараметрыОтбора.Вставить("СтруктураИерархии", ТекСтруктураИерархии);
			ПараметрыОтбора.Вставить("РазрешитьВыборФМ", Ложь);
			ПараметрыОтбора.Вставить("ТекущаяСтрока", СтарыйОР);
			
			ОткрытьФорму("Справочник.торо_ОбъектыРемонта.ФормаВыбора",ПараметрыОтбора,Элемент,,ВариантОткрытияОкна.ОтдельноеОкно,,,РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
			
		ИначеЕсли ТекДанные.ОбъектРемонта = Неопределено Тогда 
			СтандартнаяОбработка = Ложь;
			
			СписокВыбора = ЗаполнитьСписокЗначение();
			
			Если СписокВыбора.Количество() = 1 Тогда     
				
				ОбъектыРемонтаОбъектРемонтаНачалоВыбораЗавершение(СписокВыбора[0], Новый Структура("Элемент", Элемент));
				
			Иначе
				СписокВыбора.ПоказатьВыборЭлемента(Новый ОписаниеОповещения("ОбъектыРемонтаОбъектРемонтаНачалоВыбораЗавершение",ЭтотОбъект, Новый Структура("Элемент",Элемент)),НСтр("ru = 'Выбор типа данных'"));
			КонецЕсли;
			
		Иначе
			СтандартнаяОбработка = Ложь; 
			СтарыйОР = ТекДанные.ОбъектРемонта;
			ФормаВыбора = ПолучитьФорму("Справочник.торо_СписокОбъектовРегламентногоМероприятия.ФормаВыбора",, Элемент, Объект.Ссылка);
			ФормаВыбора.Открыть();
		КонецЕсли;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаНачалоВыбораЗавершение(ВыбранныйЭлемент, ДополнительныеПараметры) Экспорт
	
	Если ВыбранныйЭлемент = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Элемент = ДополнительныеПараметры.Элемент;
	ВыбЭлемент = ВыбранныйЭлемент;
	ТекДанные = Элементы.ОбъектыРемонта.ТекущиеДанные;
	
	Если ВыбЭлемент.Значение = Неопределено Тогда 
		Возврат;
	ИначеЕсли ВыбЭлемент.Значение = Тип("СправочникСсылка.торо_ОбъектыРемонта") Тогда     
		СписокДоступныхСтатусов = торо_СтатусыОРВУчете.СписокСтатусовДляПодбора();
		СтарыйОР = ТекДанные.ОбъектРемонта;
		
		ПараметрыОтбора = Новый Структура("СписокСтатусов", СписокДоступныхСтатусов);
		ПараметрыОтбора.Вставить("СтруктураИерархии",       ТекСтруктураИерархии);
		
		ОткрытьФорму("Справочник.торо_ОбъектыРемонта.ФормаВыбора",ПараметрыОтбора,Элемент,,ВариантОткрытияОкна.ОтдельноеОкно,,,РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
	Иначе    
		СтарыйОР = ТекДанные.ОбъектРемонта;
		ФормаВыбора = ПолучитьФорму("Справочник.торо_СписокОбъектовРегламентногоМероприятия.ФормаВыбора",, Элемент, Объект.Ссылка);
		ФормаВыбора.Открыть();
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	
	Если ЗначениеЗаполнено(ВыбранноеЗначение) Тогда
		Если ЭтоФункциональноеМесто(ВыбранноеЗначение) Тогда
			ТекстСообщения = НСтр("ru = 'Нельзя выбрать объект ремонта с типом функциональное место!'");
			ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
			СтандартнаяОбработка = Ложь;
		Иначе
			НайденныеСтроки = Объект.ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта",ВыбранноеЗначение));
			Если НайденныеСтроки.Количество() > 0 Тогда
				СтандартнаяОбработка = Ложь;
			КонецЕсли;
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры


&НаКлиенте
Процедура ОбъектыРемонтаПередНачаломДобавления(Элемент, Отказ, Копирование, Родитель, Группа, Параметр)
	Если Не Копирование Тогда
		ДобавлениеОР = Истина;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаПередУдалением(Элемент, Отказ)
	
	ВыделенныеСтроки = Элементы.ОбъектыРемонта.ВыделенныеСтроки;
	Для каждого УдаляемаяСтрокаИД из ВыделенныеСтроки Цикл
		УдаляемаяСтрока = Объект.ОбъектыРемонта.НайтиПоИдентификатору(УдаляемаяСтрокаИД);
		СтрокиКоэффициентов = Объект.КоэффициентыРемонтныхОсобенностей.НайтиСтроки(Новый Структура("ОбъектРемонта", УдаляемаяСтрока.ОбъектРемонта));
		Для каждого СтрокаКоэффициента из СтрокиКоэффициентов Цикл
			Объект.КоэффициентыРемонтныхОсобенностей.Удалить(СтрокаКоэффициента);
		КонецЦикла;
		
		СтрокиПодчиненных = Объект.КоэффициентыРемонтныхОсобенностейПодчиненных.НайтиСтроки(Новый Структура("РодительИерархии", УдаляемаяСтрока.ОбъектРемонта));
		Для каждого СтрокаПодчиненного из СтрокиПодчиненных Цикл
			Объект.КоэффициентыРемонтныхОсобенностейПодчиненных.Удалить(СтрокаПодчиненного);
		КонецЦикла;
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаОткрытиеФормыПодбора(ВыбранныйЭлемент, СписокПараметров=Неопределено) Экспорт
	
	Если ВыбранныйЭлемент = Неопределено Тогда
		Возврат;
	ИначеЕсли ВыбранныйЭлемент.Значение = Тип("СправочникСсылка.торо_ОбъектыРемонта") Тогда
		
		ПараметрыФормы = Новый Структура;
		ПараметрыФормы.Вставить("КлючНазначенияИспользования", "торо_УстановкаКоэффициентовРемонтныхОсобенностей");
		ПараметрыФормы.Вставить("ЗакрыватьПриВыборе", Ложь);
		ПараметрыФормы.Вставить("СтруктураИерархии", ТекСтруктураИерархии);
		ПараметрыФормы.Вставить("РазрешитьВыборФМ", Ложь);

		ОткрытьФорму("Обработка.торо_ПодборОбъектовРемонтныхРабот.Форма.ФормаПодбораОбъектовДляПланаГрафикаППР",
						ПараметрыФормы, Элементы.ОбъектыРемонта, ЭтаФорма.УникальныйИдентификатор);
						
	ИначеЕсли ВыбранныйЭлемент.Значение = Тип("СправочникСсылка.торо_СписокОбъектовРегламентногоМероприятия") Тогда
		
		ПараметрыФормы = Новый Структура;
		ПараметрыФормы.Вставить("КлючНазначенияИспользования", "торо_УстановкаКоэффициентовРемонтныхОсобенностей");
		ПараметрыФормы.Вставить("ЗакрыватьПриВыборе", Ложь);
		
		ОткрытьФорму("Обработка.торо_ПодборОбъектовРемонтныхРабот.Форма.ФормаПодбораСписковОРДляГрафикаРегламентныхМероприятий",
						ПараметрыФормы, Элементы.ОбъектыРемонта, ЭтаФорма.УникальныйИдентификатор);																		
						
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	
	Если Не ЗначениеЗаполнено(ВыбранноеЗначение) Тогда		
		Возврат;
	КонецЕсли;
	
	Если ТипЗнч(ВыбранноеЗначение) = Тип("Структура") И ВыбранноеЗначение.Свойство("ОбъектРемонта") Тогда
		
		Для каждого ВыбранныйОР Из ВыбранноеЗначение.ОбъектРемонта Цикл
			Если ТипЗнч(ВыбранныйОР) = Тип("СправочникСсылка.торо_ОбъектыРемонта") Тогда
				Если ЭтоФункциональноеМесто(ВыбранныйОР) Тогда
					ТекстСообщения = НСтр("ru = 'Нельзя выбрать объект ремонта с типом функциональное место!'");
					ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
					СтандартнаяОбработка = Ложь;
				Иначе				
					МассивОРСовпадающихСВыбранным = Объект.ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта",ВыбранныйОР));
					Если МассивОРСовпадающихСВыбранным.Количество() = 0 Тогда 
						НоваяСтрока = Объект.ОбъектыРемонта.Добавить();
						НоваяСтрока.ОбъектРемонта = ВыбранныйОР;
						Элементы.ОбъектыРемонта.ТекущаяСтрока = НоваяСтрока.ПолучитьИдентификатор();
					КонецЕсли;				
				КонецЕсли;
			КонецЕсли;
		КонецЦикла; 
		
	ИначеЕсли ТипЗнч(ВыбранноеЗначение) = Тип("Массив") Тогда 
		// Массив списков объектов регламентных мероприятий.
		// Из формы Обработка.торо_ПодборОбъектовРемонтныхРабот.Форма.ФормаПодбораСписковОРДляГрафикаРегламентныхМероприятий
		
		СтруктураОтбора = Новый Структура("ОбъектРемонта");
		
		Для каждого СписокОбъектов Из ВыбранноеЗначение Цикл
			
			СтруктураОтбора.ОбъектРемонта = СписокОбъектов;
			СписокОбъектовУжеДобавлен = Объект.ОбъектыРемонта.НайтиСтроки(СтруктураОтбора).Количество() > 0;
			
			Если Не СписокОбъектовУжеДобавлен Тогда
				НоваяСтрока = Объект.ОбъектыРемонта.Добавить();
				НоваяСтрока.ОбъектРемонта = СписокОбъектов;
				Элементы.ОбъектыРемонта.ТекущаяСтрока = НоваяСтрока.ПолучитьИдентификатор();
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;
	
КонецПроцедуры
#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыКоэффициентыРемонтныхОсобенностей
&НаКлиенте
Процедура КоэффициентыРемонтныхОсобенностейПриАктивизацииСтроки(Элемент)
	ТекДанные = Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные;
	Если НЕ ТекДанные = Неопределено Тогда
		флагИспользование = Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные.Использование;
		Элементы.КоэффициентыРемонтныхОсобенностейЗначениеКоэффициента.Доступность = флагИспользование;
		
		ТекВидКоэф = ТекДанные.ВидКоэффициента;
		ТекДатаНачИсп = ТекДанные.ДатаНачалаИспользования;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыРемонтныхОсобенностейПриНачалеРедактирования(Элемент, НоваяСтрока, Копирование)
	Если НоваяСтрока и НЕ Копирование Тогда
		Если НоваяСтрока И Элементы.ОбъектыРемонта.ТекущиеДанные <> Неопределено Тогда
			Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные.ОбъектРемонта = Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта;
		КонецЕсли;
		Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные.Использование = Истина;
		КоэффициентыРемонтныхОсобенностейИспользованиеПриИзменении(Элементы.КоэффициентыРемонтныхОсобенностей);
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыРемонтныхОсобенностейПередОкончаниемРедактирования(Элемент, НоваяСтрока, ОтменаРедактирования, Отказ)
	ТекСтр = Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные;
	Если НЕ ОтменаРедактирования Тогда
		Если НоваяСтрока И Элементы.ОбъектыРемонта.ТекущиеДанные <> Неопределено Тогда
			ТекСтр.ОбъектРемонта = Элементы.ОбъектыРемонта.ТекущиеДанные.ОбъектРемонта;
		КонецЕсли;
		МассивСтрок = Объект.КоэффициентыРемонтныхОсобенностей.НайтиСтроки(Новый Структура("ОбъектРемонта, ВидКоэффициента, ДатаНачалаИспользования",ТекСтр.ОбъектРемонта, ТекСтр.ВидКоэффициента, ТекСтр.ДатаНачалаИспользования));
		Если МассивСтрок.Количество() > 1 Тогда
			ШаблонСообщения = НСтр("ru = 'Для объекта ремонта %1 уже установлен вид коэффициента %2 (с %3)!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, ТекСтр.ОбъектРемонта, ТекСтр.ВидКоэффициента, ТекСтр.ДатаНачалаИспользования);
			ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения,,,,Отказ);
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыРемонтныхОсобенностейРаспространятьНаПодчиненныхПриИзменении(Элемент)
	
	ТекДанные = Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные;
	Если НЕ ЗначениеЗаполнено(ТекДанные.СтруктураИерархии) Тогда
		ТекДанные.РаспространятьНаПодчиненных= НЕ ТекДанные.РаспространятьНаПодчиненных;
		ТекстСообщения = НСтр("ru = 'Не заполнено значение структуры иерархии.'");
		ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
		Возврат;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыРемонтныхОсобенностейИспользованиеПриИзменении(Элемент)
	
	флагИспользование = Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные.Использование;
	Если флагИспользование Тогда 
		Элементы.КоэффициентыРемонтныхОсобенностейЗначениеКоэффициента.Доступность = Истина;
	Иначе
		Элементы.КоэффициентыРемонтныхОсобенностейЗначениеКоэффициента.Доступность = Ложь;
		Элементы.КоэффициентыРемонтныхОсобенностей.ТекущиеДанные.ЗначениеКоэффициента = ПредопределенноеЗначение("Справочник.торо_ВидыКоэффициентовРемонтныхОсобенностей.ПустаяСсылка");
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

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

// СтандартныеПодсистемы.Свойства
&НаКлиенте
Процедура Подключаемый_СвойстваВыполнитьКоманду(ЭлементИлиКоманда, НавигационнаяСсылка = Неопределено, СтандартнаяОбработка = Неопределено)
	УправлениеСвойствамиКлиент.ВыполнитьКоманду(ЭтотОбъект, ЭлементИлиКоманда, СтандартнаяОбработка);
КонецПроцедуры
// Конец СтандартныеПодсистемы.Свойства

&НаКлиенте
Процедура Подбор(Команда)
	СписокВыбора = ЗаполнитьСписокЗначение();
	Если СписокВыбора.Количество() = 1 Тогда     		
		ОбъектыРемонтаОбъектРемонтаОткрытиеФормыПодбора(ЭтотОбъект);		
	Иначе
		СписокВыбора.ПоказатьВыборЭлемента(Новый ОписаниеОповещения("ОбъектыРемонтаОбъектРемонтаОткрытиеФормыПодбора",ЭтотОбъект),НСтр("ru = 'Выбор типа данных'"));
	КонецЕсли;
КонецПроцедуры
#КонецОбласти

#Область СлужебныеПроцедурыИФункции

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

&НаСервереБезКонтекста
Функция ЗаполнитьСписокЗначение()
	
	ФОИспользоватьРегламентныеМероприятия = ПолучитьФункциональнуюОпцию("торо_ИспользоватьРегламентныеМероприятия");
	
	СписокВыбора = Новый СписокЗначений;
	
	Для Каждого ТипРеквизита Из Метаданные.Документы.торо_УстановкаКоэффициентовРемонтныхОсобенностей.ТабличныеЧасти.ОбъектыРемонта.Реквизиты.ОбъектРемонта.Тип.Типы() Цикл
		Если Не ФОИспользоватьРегламентныеМероприятия И ТипРеквизита = Тип("СправочникСсылка.торо_СписокОбъектовРегламентногоМероприятия") Тогда
	
			Продолжить;
		
		КонецЕсли;
	
		СписокВыбора.Добавить(ТипРеквизита);
	КонецЦикла;
	
	Возврат СписокВыбора;
	
КонецФункции

&НаКлиенте
Процедура ОбъектыРемонтаПриНачалеРедактирования(Элемент, НоваяСтрока, Копирование)
	Если Копирование Тогда
		текДанные = Элементы.ОбъектыРемонта.ТекущиеДанные;
		текДанные.ОбъектРемонта = ПредопределенноеЗначение("Справочник.торо_ОбъектыРемонта.ПустаяСсылка");
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаОчистка(Элемент, СтандартнаяОбработка)
	Если Элементы.ОбъектыРемонта.ТекущиеДанные <> Неопределено Тогда
		ТекущийОР = Элементы.ОбъектыРемонта.ТекущиеДанные;
		
		СтрокиКоэффициентов = Объект.КоэффициентыРемонтныхОсобенностей.НайтиСтроки(Новый Структура("ОбъектРемонта", ТекущийОР.ОбъектРемонта));
		Для каждого СтрокаКоэффициента из СтрокиКоэффициентов Цикл
			Объект.КоэффициентыРемонтныхОсобенностей.Удалить(СтрокаКоэффициента);
		КонецЦикла;
		
		СтрокиПодчиненныхНаУдаление = Объект.КоэффициентыРемонтныхОсобенностейПодчиненных.НайтиСтроки(Новый Структура("РодительИерархии", ТекущийОР.ОбъектРемонта));
		
		Для Каждого СтрокаПодчиненного Из СтрокиПодчиненныхНаУдаление Цикл 
			Объект.КоэффициентыРемонтныхОсобенностейПодчиненных.Удалить(СтрокаПодчиненного);
		КонецЦикла;
	
	КонецЕсли; 

КонецПроцедуры

&НаКлиенте
Процедура КомментарийНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
		
	СтандартнаяОбработка = Ложь;
	ОбщегоНазначенияКлиент.ПоказатьФормуРедактированияКомментария(
		Элемент.ТекстРедактирования, ЭтотОбъект, "Объект.Комментарий");	
	
	КонецПроцедуры
	
&НаСервере	
Функция ЭтоФункциональноеМесто(ДобавляемыйОбъект)
	
	Если ТипЗнч(ДобавляемыйОбъект) = Тип("СправочникСсылка.торо_ОбъектыРемонта")
		И Константы.торо_ИспользоватьФункциональныеМеста.Получить()
		И ДобавляемыйОбъект.ТипОбъекта = Перечисления.торо_ТипыОбъектовRCM.ФункциональноеМесто Тогда
		Возврат Истина;
	КонецЕсли;
	
	Возврат Ложь;
	
КонецФункции

#КонецОбласти
