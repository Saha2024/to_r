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
	// Конец СтандартныеПодсистемы.ПодключаемыеКомандыи
	
	// СтандартныеПодсистемы.ВерсионированиеОбъектов
	ВерсионированиеОбъектов.ПриСозданииНаСервере(ЭтаФорма);
	// Конец СтандартныеПодсистемы.ВерсионированиеОбъектов
	
	мОтображатьПоложение = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить(
			"НастройкиТОиР",
			"ПоказыватьПоложениеОР",
			Истина);
			
	Элементы.ОбъектыРемонтаОтображатьПоложение.Пометка = мОтображатьПоложение;
	Элементы.ОбъектыРемонтаПоложение.Видимость = мОтображатьПоложение;
	
	ТекСтруктураИерархии = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ОсновнаяСтруктураИерархии");
	
	ЗаполнитьСписокОбъектовРемонта();
	ЗаполнитьГраничныеЗначенияПоказателей();
	ЗаполнитьПараметрыОформленияЗначенийКПОР();
	
	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
	
	// Установить свойства элементов формы в зависимости от доступных прав
	торо_СобытияФорм.УстановитьСвойствоЭлементовФормыОтПрав(ЭтотОбъект);

	Если мОтображатьПоложение Тогда	
		ЗаполнитьПоложенияОР();
	КонецЕсли;
	
	торо_УправлениеДоступом.УстановитьДоступностьРедактированияПоОграничениюДоступаДляОР(ЭтаФорма, Объект.Показатели.Выгрузить());
	
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
	
	// СтандартныеПодсистемы.ДатыЗапретаИзменения
	ДатыЗапретаИзменения.ОбъектПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.ДатыЗапретаИзменения	
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	// СтандартныеПодсистемы.Свойства
    УправлениеСвойствамиКлиент.ПослеЗагрузкиДополнительныхРеквизитов(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиент.НачатьОбновлениеКоманд(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	Если НЕ ЗначениеЗаполнено(Объект.Ссылка) 
		И ЗначениеЗаполнено(Объект.ДокументОснование) Тогда
		ПодключитьОбработчикОжидания("ПодборОРОбработчикОжидания",0.5,Истина);
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура ОбработкаПроверкиЗаполненияНаСервере(Отказ, ПроверяемыеРеквизиты)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ОбработкаПроверкиЗаполнения(ЭтотОбъект, Отказ, ПроверяемыеРеквизиты);
	// Конец СтандартныеПодсистемы.Свойства
	Для Каждого СтрокаТЧ Из Объект.Показатели Цикл
		Отбор = Новый Структура;
		Отбор.Вставить("ОбъектРемонта",СтрокаТЧ.ОбъектРемонта);
		Отбор.Вставить("Показатель",СтрокаТЧ.Показатель);
		Отбор.Вставить("ДатаКонтроля", СтрокаТЧ.ДатаКонтроля);
		
		Строки = Объект.Показатели.НайтиСтроки(Отбор);
		КоличествоСтрок = Строки.Количество(); 
		Если СтрокаТЧ.ДатаКонтроля = Объект.Дата Тогда
			Отбор.Вставить("ДатаКонтроля", Дата(1,1,1,0,0,0));
			Строки2 = Объект.Показатели.НайтиСтроки(Отбор);
			КоличествоСтрок = КоличествоСтрок + Строки2.Количество();
		КонецЕсли; 
		
		Если КоличествоСтрок > 1 Тогда
			ШаблонСообщения = НСтр("ru = 'Обнаружены показатели ""(%1)"", значения на указанную дату которых повторяются!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаТЧ.Показатель);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,"Объект.Показатели",,Отказ);
			Возврат;
		КонецЕсли;
		
		Если НЕ ЗначениеЗаполнено(СтрокаТЧ.Показатель) Тогда
			ТекстСообщения = НСтр("ru = 'Обнаружено незаполненное поле ""Показатель"" в таблице ""Контролируемые показатели""!'");
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,"Объект.Показатели",,Отказ);
			Возврат;
		КонецЕсли;
		
		Если НЕ ЗначениеЗаполнено(СтрокаТЧ.Значение) И СтрокаТЧ.Значение <> 0 Тогда
			ШаблонСообщения = НСтр("ru = 'Обнаружено незаполненное поле ""Количество"" для показателя ""(%1)"" в таблице ""Контролируемые показатели""!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, СтрокаТЧ.Показатель);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,"Объект.Показатели",,Отказ);
			Возврат;  
		КонецЕсли;
	КонецЦикла;
КонецПроцедуры

&НаСервере
Процедура ПередЗаписьюНаСервере(Отказ, ТекущийОбъект, ПараметрыЗаписи)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ПередЗаписьюНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	// Сохранить изменение порядка ОР стрелочками
	КопияПоказателейДляПроверки = ТекущийОбъект.Показатели.Выгрузить();
	КопияПоказателейДляПроверки.Свернуть("ОбъектРемонта");
	НужнаПерестановка = Ложь;
	КоличествоОР = Мин(ОбъектыРемонта.Количество(), КопияПоказателейДляПроверки.Количество());
	
   Для Индекс = 0 по КоличествоОР - 1 Цикл
		Если КопияПоказателейДляПроверки[Индекс].ОбъектРемонта <> ОбъектыРемонта[Индекс].ОбъектРемонта Тогда
			НужнаПерестановка = Истина;
			Прервать;
		КонецЕсли;
	КонецЦикла;
	
	Если НужнаПерестановка Тогда
		КопияПоказателей = ТекущийОбъект.Показатели.Выгрузить();
		ТекущийОбъект.Показатели.Очистить();
		КопияПоказателей.Колонки.Добавить("Сохранено", Новый ОписаниеТипов("Булево"));
		
		Для каждого СтрокаОР из ОбъектыРемонта Цикл
			МассивСтрок = КопияПоказателей.НайтиСтроки(Новый Структура("ОбъектРемонта", СтрокаОР.ОбъектРемонта));
			Для каждого ПереносимаяСтрока из МассивСтрок Цикл
				Если НЕ ПереносимаяСтрока.Сохранено Тогда
					НовСтр = ТекущийОбъект.Показатели.Добавить();
					ЗаполнитьЗначенияСвойств(НовСтр,ПереносимаяСтрока);
					ПереносимаяСтрока.Сохранено = Истина;
				КонецЕсли;
			КонецЦикла;
		КонецЦикла;
		
		МассивСтрок = КопияПоказателей.НайтиСтроки(Новый Структура("Сохранено", Ложь));
		Для каждого ПереносимаяСтрока из МассивСтрок Цикл
			НовСтр = ТекущийОбъект.Показатели.Добавить();
			ЗаполнитьЗначенияСвойств(НовСтр,ПереносимаяСтрока);
			ПереносимаяСтрока.Сохранено = Истина;
		КонецЦикла;
		
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура ПослеЗаписиНаСервере(ТекущийОбъект, ПараметрыЗаписи)
	
	// СтандартныеПодсистемы.УправлениеДоступом
	УправлениеДоступом.ПослеЗаписиНаСервере(ЭтотОбъект, ТекущийОбъект, ПараметрыЗаписи);
	// Конец СтандартныеПодсистемы.УправлениеДоступом

	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
	
	ЗаполнитьГраничныеЗначенияПоказателей();
	ЗаполнитьПараметрыОформленияЗначенийКПОР();
	
	// Ремонты по состоянию
	ДокументыВДКОткрытию.Очистить();
	Если ПараметрыЗаписи.РежимЗаписи = РежимЗаписиДокумента.Проведение
		И ПолучитьФункциональнуюОпцию("торо_ИспользоватьРемонтыПоСостоянию") Тогда
		торо_РемонтыПоСостоянию.СформироватьУведомленияИДефекты(ТекущийОбъект,ДокументыВДКОткрытию);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПослеЗаписи(ПараметрыЗаписи)
	
	Если ЗначениеЗаполнено(Объект.ДокументОснование) 
		И (ТипЗнч(Объект.ДокументОснование) = Тип("ДокументСсылка.торо_АктОВыполненииЭтапаРабот") 
		ИЛИ ТипЗнч(Объект.ДокументОснование) = Тип("ДокументСсылка.торо_АктОВыполненииРегламентногоМероприятия")) Тогда
		Оповестить("ЗаписанПоказательНаработкаСостояние");
	КонецЕсли;

	Для Каждого СтруктураДокументаВДКОткрытию Из ДокументыВДКОткрытию Цикл
		
		Если СтруктураДокументаВДКОткрытию.Значение.Свойство("СсылкаНаДокумент") Тогда
			
			ПараметрыФормы = Новый Структура("Ключ", СтруктураДокументаВДКОткрытию.Значение.СсылкаНаДокумент);
			ОткрытьФорму("Документ.торо_ВыявленныеДефекты.ФормаОбъекта", ПараметрыФормы);
			
		Иначе
			
			ЗначенияЗаполнения = Новый Структура("ДанныеДокументаПриСозданииПоСостоянию", СтруктураДокументаВДКОткрытию.Значение);
			ПараметрыФормы = Новый Структура("ЗначенияЗаполнения", ЗначенияЗаполнения);
			
			КлючУникальностиФормы = Строка(СтруктураДокументаВДКОткрытию.Значение.СтруктураДокументаОснования.ДокументИсточник.УникальныйИдентификатор())
											+СтруктураДокументаВДКОткрытию.Значение.СтруктураДокументаОснования.ОбъектРемонта.УникальныйИдентификатор()
											+СтруктураДокументаВДКОткрытию.Значение.СтруктураДокументаОснования.Показатель.УникальныйИдентификатор();
			
			ОткрытьФорму("Документ.торо_ВыявленныеДефекты.ФормаОбъекта", ПараметрыФормы,,КлючУникальностиФормы);
			
		КонецЕсли;
		
	КонецЦикла;
	
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

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ДатаПриИзменении(Элемент)
	
	Для каждого СтрокаТЧ Из Объект.Показатели Цикл
		Отбор=Новый Структура;
		Отбор.Вставить("ОбъектРемонта",СтрокаТЧ.ОбъектРемонта);
		Отбор.Вставить("Показатель",СтрокаТЧ.Показатель);
		Отбор.Вставить("ДатаКонтроля",СтрокаТЧ.ДатаКонтроля);
		
		Строки = Объект.Показатели.НайтиСтроки(Отбор);
		КоличествоСтрок = Строки.Количество(); 
		Если СтрокаТЧ.ДатаКонтроля = Объект.Дата Тогда
			Отбор.Вставить("ДатаКонтроля", Дата(1,1,1,0,0,0));
			Строки2 = Объект.Показатели.НайтиСтроки(Отбор);
			КоличествоСтрок = КоличествоСтрок + Строки2.Количество();
		КонецЕсли; 
		
		Если КоличествоСтрок > 1 Тогда
			ШаблонСообщения = НСтр("ru = 'В документе для %1 значения по %2 на эту дату уже введены.'");
			ТекстСообщения =  СтрШаблон(ШаблонСообщения, СтрокаТЧ.ОбъектРемонта, СтрокаТЧ.Показатель);
			ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
		КонецЕсли;
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура КомментарийНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	ОбщегоНазначенияКлиент.ПоказатьФормуРедактированияКомментария(
		Элемент.ТекстРедактирования, ЭтотОбъект, "Объект.Комментарий");	
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыОбъектыРемонта

&НаКлиенте
Процедура ОбъектыРемонтаПриАктивизацииСтроки(Элемент)
	
	ТекущаяСтрока = Элементы.ОбъектыРемонта.ТекущиеДанные;
	СтрокаВыбрана = Не (ТекущаяСтрока = Неопределено);
	
	// Отбор строк по объектам ремонта
	Если СтрокаВыбрана Тогда
		Элементы.Показатели.ОтборСтрок = Неопределено;
		Элементы.Показатели.ОтборСтрок = Новый ФиксированнаяСтруктура("ОбъектРемонта", ТекущаяСтрока.ОбъектРемонта);					
		ТекущийОР = ТекущаяСтрока.ОбъектРемонта;
		ОР = ТекущаяСтрока.ОбъектРемонта;
	КонецЕсли;
	
	Элементы.Показатели.Доступность = СтрокаВыбрана;	
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаПередНачаломДобавления(Элемент, Отказ, Копирование, Родитель, Группа)
	
	Отказ = Истина;
	
	Если Не Копирование Тогда                                                                                
		СписокДоступныхСтатусов = торо_СтатусыОРВУчете.СписокСтатусовДляПодбора();
				
		ПараметрыФормы = Новый Структура("СписокСтатусов, СтруктураИерархии, ЕстьИзмеряемыеПоказатели", СписокДоступныхСтатусов, ТекСтруктураИерархии, Истина);
		ОткрытьФорму("Справочник.торо_ОбъектыРемонта.ФормаВыбора", ПараметрыФормы, Элемент);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаПередУдалением(Элемент, Отказ)
	
	ТекущаяСтрОР = Элементы.ОбъектыРемонта.ТекущиеДанные;
	Если НЕ ТекущаяСтрОР = Неопределено Тогда 
		Строки = Объект.Показатели.НайтиСтроки(Новый Структура("ОбъектРемонта", ТекущаяСтрОР.ОбъектРемонта));
		Для Каждого СтрокаТЧ Из Строки Цикл 
			Объект.Показатели.Удалить(СтрокаТЧ);
		КонецЦикла;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	
	Если ВыбранноеЗначение = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	ЗаполнитьПоказателиНаСервере(ВыбранноеЗначение);	
	
	Если мОтображатьПоложение Тогда	
		ЗаполнитьПоложенияОР();
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаПриИзменении(Элемент)
	
	Строки = Объект.Показатели.НайтиСтроки(Новый Структура("ОбъектРемонта", ТекущийОР));
	Для Каждого СтрокаТЧ Из Строки Цикл 
		Объект.Показатели.Удалить(СтрокаТЧ);
	КонецЦикла;
	
	Если мОтображатьПоложение Тогда
		ЗаполнитьПоложениеОРВстроке(Элементы.ОбъектыРемонта.ТекущиеДанные);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаПередОкончаниемРедактирования(Элемент, НоваяСтрока, ОтменаРедактирования, Отказ)
	
	ТекСтрока = Элементы.ОбъектыРемонта.ТекущиеДанные;
	Если НЕ (НоваяСтрока И ОтменаРедактирования) Тогда
		Если НЕ ЗначениеЗаполнено(ТекСтрока.ОбъектРемонта) Тогда 
			ТекстСообщения = НСтр("ru = 'Не заполнен объект ремонта!'");
			ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения,,,,Отказ);
			Возврат;
		КонецЕсли;
		
		Строки = ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта", ТекСтрока.ОбъектРемонта));
		Если Строки.Количество() > 1 Тогда 
			ТекстСообщения = НСтр("ru = 'Данный объект ремонта уже присутствует в табличном поле!'");
			ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения,,,,Отказ);
			Возврат;
		КонецЕсли;
		
		Элементы.Показатели.ОтборСтрок = Неопределено;
		Элементы.Показатели.ОтборСтрок = Новый ФиксированнаяСтруктура("ОбъектРемонта", ТекСтрока.ОбъектРемонта);					
		ТекущийОР = ТекСтрока.ОбъектРемонта;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	СписокДоступныхСтатусов = торо_СтатусыОРВУчете.СписокСтатусовДляПодбора();
	
	ПараметрыОтбора = Новый Структура("СписокСтатусов, СтруктураИерархии", СписокДоступныхСтатусов, ТекСтруктураИерархии);
	
	ТекущиеДанные = Элементы.ОбъектыРемонта.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено Тогда
		ПараметрыОтбора.Вставить("ТекущаяСтрока", ТекущиеДанные.ОбъектРемонта);
	КонецЕсли;
	
	ОткрытьФорму("Справочник.торо_ОбъектыРемонта.ФормаВыбора", ПараметрыОтбора, Элемент);
	
КонецПроцедуры

&НаКлиенте
Процедура ОбъектыРемонтаОбъектРемонтаОткрытие(Элемент, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	ТекущиеДанные = Элементы.ОбъектыРемонта.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Если ЗначениеЗаполнено(ТекущиеДанные.ОбъектРемонта) Тогда
		ПараметрыФормы = Новый Структура("Ключ, СтруктураИерархии", ТекущиеДанные.ОбъектРемонта, ТекСтруктураИерархии);
		ОткрытьФорму("Справочник.торо_ОбъектыРемонта.ФормаОбъекта", ПараметрыФормы);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыПоказатели

&НаКлиенте
Процедура ПоказателиПередНачаломДобавления(Элемент, Отказ, Копирование, Родитель, Группа)
	
	ТекСтрокаОР = Элементы.ОбъектыРемонта.ТекущиеДанные;
	Если ТекСтрокаОР = Неопределено Тогда
		
		Отказ = Истина;
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Укажите оборудование для которого фиксируется значение контролируемых показателей!'"), 20);
		               		
	ИначеЕсли Не ЗначениеЗаполнено(ТекСтрокаОР.ОбъектРемонта) Тогда
		
		Отказ = Истина;
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Не заполнено оборудование для которого фиксируется значение контролируемых показателей!'"), 20);
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПриНачалеРедактирования(Элемент, НоваяСтрока, Копирование)
	
	ТекДанные = Элемент.ТекущиеДанные;
	
	Если НоваяСтрока Тогда
		ТекСтрокаОР = Элементы.ОбъектыРемонта.ТекущиеДанные;
		ТекДанные.ОбъектРемонта = ТекСтрокаОР.ОбъектРемонта;
	КонецЕсли;
	
	Если НЕ НоваяСтрока И НЕ Копирование Тогда
		ЗаполнитьЗначенияСвойств(ПоказательДоРедактирования,ТекДанные);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПередОкончаниемРедактирования(Элемент, НоваяСтрока, ОтменаРедактирования, Отказ)	
	Если НЕ НоваяСтрока И ОтменаРедактирования Тогда
		СтрокаТЧ = Элемент.ТекущиеДанные;
		ЗаполнитьЗначенияСвойств(СтрокаТЧ,ПоказательДоРедактирования);
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПоказательНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	ТекСтрока = Элементы.Показатели.ТекущиеДанные;
	
	Если ТекСтрока <> Неопределено И Не ЗначениеЗаполнено(ТекСтрока.ОбъектРемонта) Тогда
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Для выбора значения контролируемого показателя необходимо указать объект ремонта!'"), 60);
		Возврат;
	КонецЕсли;
	
	СписокОтбора = ПолучитьСписокПоказателейОбъектаРемонта(ТекСтрока.ОбъектРемонта);
	ПараметрыОтбора = Новый Структура("Ссылка", СписокОтбора);
	ПараметрыФормы = Новый Структура("Отбор", ПараметрыОтбора);
	
	ОткрытьФорму("ПланВидовХарактеристик.торо_ИзмеряемыеПоказателиОбъектовРемонта.ФормаВыбора", ПараметрыФормы, Элемент, УникальныйИдентификатор);  
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПоказательОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	ТекСтрока = Элементы.Показатели.ТекущиеДанные;
	Если ТекСтрока <> Неопределено Тогда
		СписокОтбора = ПолучитьСписокПоказателейОбъектаРемонта(ТекСтрока.ОбъектРемонта);
		СтрокаПоказателя = СписокОтбора.НайтиПоЗначению(ВыбранноеЗначение);
		Если СтрокаПоказателя = Неопределено Тогда
			СтандартнаяОбработка = Ложь;
		КонецЕсли; 
	Иначе
		СтандартнаяОбработка = Ложь;
	КонецЕсли; 
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПоказательАвтоПодбор(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, Ожидание, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ТекСтрока = Элементы.Показатели.ТекущиеДанные;
	ДанныеВыбора = ПолучитьСписокПоказателейОбъектаРемонта(ТекСтрока.ОбъектРемонта);
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПоказательОкончаниеВводаТекста(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, СтандартнаяОбработка)
	ТекСтрока = Элементы.Показатели.ТекущиеДанные;
	СписокОтбора = ПолучитьСписокПоказателейОбъектаРемонта(ТекСтрока.ОбъектРемонта);
	ПараметрыПолученияДанных.Отбор.Вставить("Ссылка", СписокОтбора);
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиПоказательПриИзменении(Элемент)
	
	ТекСтрока = Элементы.Показатели.ТекущиеДанные;
	СтруктураВозврата = ПолучитьЗначенияИзмеряемыхПоказателейНаСервере(ТекСтрока.ОбъектРемонта, ТекСтрока.Показатель);
	ЗаполнитьЗначенияСвойств(ТекСтрока, СтруктураВозврата,,"ОбъектРемонта");
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказателиЗначениеПриИзменении(Элемент)
	ТекущиеДанные = Элементы.Показатели.ТекущиеДанные;
	Если ТекущиеДанные = Неопределено Тогда
	    Возврат;
	КонецЕсли;
	
	Если ЗначениеЗаполнено(ТекущиеДанные.Значение) И ТипЗнч(ТекущиеДанные.Значение) = Тип("СправочникСсылка.торо_ЗначенияИзмеряемыхПоказателейОбъектовРемонта") Тогда
	    ТекущиеДанные.ЗначениеКПОР = Истина;
		ТекущиеДанные.ВыделятьКакНедопустимоеЗначение = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Значение, "ВыделятьКакНедопустимоеЗначение");
		ТекущиеДанные.ВыделятьКакКритическоеЗначение = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.Значение, "ВыделятьКакКритическоеЗначение");
	Иначе
		ТекущиеДанные.ЗначениеКПОР = Ложь;
		ТекущиеДанные.ВыделятьКакНедопустимоеЗначение = Ложь;
		ТекущиеДанные.ВыделятьКакКритическоеЗначение = Ложь;
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
Процедура ОтображатьПоложение(Команда)
	
	Кнопка = Элементы.ОбъектыРемонтаОтображатьПоложение;
	Кнопка.Пометка = НЕ Кнопка.Пометка;	
	Элементы.ОбъектыРемонтаПоложение.Видимость = Кнопка.Пометка;
	Если Кнопка.Пометка Тогда
		ЗаполнитьПоложенияОР(); 
	КонецЕсли;
	
	мОтображатьПоложение = Кнопка.Пометка;
	
КонецПроцедуры

&НаКлиенте
Процедура НастройкаИерархии(Команда)
		
	ОписаниеОповещения = Новый ОписаниеОповещения("НастройкаИерархииЗавершение", ЭтотОбъект);
	ПараметрыОткрытия = Новый Структура("ТекущаяСтрока", ТекСтруктураИерархии);
	ОткрытьФорму("Справочник.торо_СтруктурыОР.ФормаВыбора", ПараметрыОткрытия, ЭтаФорма, ЭтаФорма.УникальныйИдентификатор,,,ОписаниеОповещения, РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
	
КонецПроцедуры

&НаКлиенте
Процедура Подбор(Команда)
	
	Если ЗначениеЗаполнено(Объект.ДокументОснование) Тогда 		
		
		ОткрытьФормуПодбораОРИзОснования(Элементы.ОбъектыРемонта);
		
	Иначе       		
		
		ПараметрыФормы = Новый Структура("КлючНазначенияИспользования, ЗакрыватьПриВыборе, СтруктураИерархии, ЕстьИзмеряемыеПоказатели", "торо_УчетКонтролируемыхПоказателей", Ложь,ТекСтруктураИерархии, Истина);
		//ЗначениеОтбора = Новый Структура("ЕстьИзмеряемыеПоказатели", Истина);
		//ПараметрыФормы.Вставить("Отбор", ЗначениеОтбора);
		ОткрытьФорму("Обработка.торо_ПодборОбъектовРемонтныхРабот.Форма.ФормаПодбораОбъектовДляПланаГрафикаППР",
			ПараметрыФормы, Элементы.ОбъектыРемонта, ЭтаФорма.УникальныйИдентификатор);
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ЗаполнитьКонтролируемыеПоказатели(Команда)
	
	ТекущиеДанные = Элементы.ОбъектыРемонта.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено Тогда
		ЗаполнитьКонтролируемыеПоказателиНаСервере(ТекущиеДанные.ОбъектРемонта);
	Иначе
		ТекстСообщения = НСтр("ru = 'Заполнение контролируемых показателей возможно только для объекта ремонта.'");
		ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
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

&НаСервере
Процедура ЗаполнитьПоложенияОР()
	
	МассивОР = Новый Массив;
	Для каждого Стр Из ОбъектыРемонта Цикл
		МассивОР.Добавить(Стр.ОбъектРемонта);		
	КонецЦикла; 
	
	СтруктураПоложений = торо_РаботаСИерархией20.ПолучитьТекущихРодителейВИерархии(МассивОР, ТекСтруктураИерархии,,Истина,Истина);
	
	Для каждого Строка из ОбъектыРемонта Цикл
		Строка.Положение = СтруктураПоложений[Строка.ОбъектРемонта];
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура ЗаполнитьПоложениеОРВстроке(ТекущиеДанные)
	
	Если ТекущиеДанные = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	ТекущиеДанные.Положение = торо_РаботаСИерархией20.ПолучитьТекущегоРодителяВИерархии(ТекущиеДанные.ОбъектРемонта, ТекСтруктураИерархии,,Истина,Истина);
	
КонецПроцедуры

&НаКлиенте
Процедура НастройкаИерархииЗавершение(РезультатЗакрытия, ДополнительныеПараметры) Экспорт
	
	Если РезультатЗакрытия <> Неопределено Тогда
		ТекСтруктураИерархии = РезультатЗакрытия;
	КонецЕсли;
	
	Если мОтображатьПоложение Тогда	
		ЗаполнитьПоложенияОР();
	КонецЕсли;
	
КонецПроцедуры

// Процедура выполняет первоначальное заполнение списка объектов ремонта.
&НаСервере
Процедура ЗаполнитьСписокОбъектовРемонта()
	
	Объекты = Объект.Показатели.Выгрузить();
	Объекты.Свернуть("ОбъектРемонта");
	ОбъектыРемонта.Загрузить(Объекты);
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьГраничныеЗначенияПоказателей()
	
	табПоказатели =  Объект.Показатели.Выгрузить();
	СписокОбъектовРемонта = табПоказатели.ВыгрузитьКолонку("ОбъектРемонта");
	СписокПоказателей = табПоказатели.ВыгрузитьКолонку("Показатель");
	МассивПоказателей = ПланыВидовХарактеристик.торо_ИзмеряемыеПоказателиОбъектовРемонта.ПолучитьСтруктуруИзмеряемыхПоказателейОбъектовРемонта(СписокОбъектовРемонта,СписокПоказателей,Истина);
	
	СтруктураПоиска = Новый Структура("ОбъектРемонта, Показатель");
	Для каждого СтруктураВозврата из МассивПоказателей Цикл
		ЗаполнитьЗначенияСвойств(СтруктураПоиска, СтруктураВозврата);
		СтруктураПоиска.ОбъектРемонта = СтруктураВозврата.ИсходныйОР;
		НайденныеСтроки = Объект.Показатели.НайтиСтроки(СтруктураПоиска);
		Для каждого НайденнаяСтрока из НайденныеСтроки Цикл
			ЗаполнитьЗначенияСвойств(НайденнаяСтрока, СтруктураВозврата);
			НайденнаяСтрока.ОбъектРемонта = СтруктураВозврата.ИсходныйОР;
		КонецЦикла;
	КонецЦикла;
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьПараметрыОформленияЗначенийКПОР()
	Для каждого СтрокаПоказателя Из Объект.Показатели Цикл
		Если Не ТипЗнч(СтрокаПоказателя.Значение) = Тип("СправочникСсылка.торо_ЗначенияИзмеряемыхПоказателейОбъектовРемонта") Тогда
		    Продолжить;
		КонецЕсли;
		
		СтрокаПоказателя.ЗначениеКПОР = Истина;
		СтрокаПоказателя.ВыделятьКакНедопустимоеЗначение = СтрокаПоказателя.Значение.ВыделятьКакНедопустимоеЗначение;
		СтрокаПоказателя.ВыделятьКакКритическоеЗначение = СтрокаПоказателя.Значение.ВыделятьКакКритическоеЗначение;
	КонецЦикла;
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьПоказателиНаСервере(ВыбранноеЗначение)
	
	Если ТипЗнч(ВыбранноеЗначение) = Тип("Массив") Тогда
		
		Для Каждого Элем Из ВыбранноеЗначение Цикл
			МассивСтрок = ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта", Элем));
			Если МассивСтрок.Количество() = 0 Тогда
				НовСтрокаОРР = ОбъектыРемонта.Добавить();
				НовСтрокаОРР.ОбъектРемонта = Элем; 
			КонецЕсли;			
		КонецЦикла;
		
	ИначеЕсли ТипЗнч(ВыбранноеЗначение) = Тип("СписокЗначений") Тогда
		
		Для Каждого Элем Из ВыбранноеЗначение Цикл
			МассивСтрок = ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта", Элем.Значение));
			Если МассивСтрок.Количество() = 0 Тогда
				НовСтрокаОРР = ОбъектыРемонта.Добавить();
				НовСтрокаОРР.ОбъектРемонта = Элем.Значение; 
			КонецЕсли;			
		КонецЦикла;
	
	ИначеЕсли ТипЗнч(ВыбранноеЗначение) = Тип("Структура") Тогда
		
		Если ВыбранноеЗначение.Свойство("ОбъектРемонта") И ЗначениеЗаполнено(ВыбранноеЗначение.ОбъектРемонта) Тогда
			МассивСтрок = ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта", ВыбранноеЗначение.ОбъектРемонта));
			Если МассивСтрок.Количество() = 0 Тогда
				НовСтрокаОРР = ОбъектыРемонта.Добавить();
				НовСтрокаОРР.ОбъектРемонта = ВыбранноеЗначение.ОбъектРемонта;
				Элементы.ОбъектыРемонта.ТекущаяСтрока = НовСтрокаОРР.ПолучитьИдентификатор();
			КонецЕсли;
			
			Если ВыбранноеЗначение.Свойство("КонтролируемыеПоказатели") Тогда
				Для каждого СтрокаПоказатель Из ВыбранноеЗначение.КонтролируемыеПоказатели Цикл
					НайденныеСтроки = Объект.Показатели.НайтиСтроки(Новый Структура("ОбъектРемонта, Показатель", ВыбранноеЗначение.ОбъектРемонта, СтрокаПоказатель.Показатель));
					Если НайденныеСтроки.Количество() = 0 Тогда
						СтрокаТЧ = Объект.Показатели.Добавить();
						ЗаполнитьЗначенияСвойств(СтрокаТЧ, СтрокаПоказатель);
						СтрокаТЧ.ОбъектРемонта = ВыбранноеЗначение.ОбъектРемонта;
					КонецЕсли;
				КонецЦикла;
			КонецЕсли;
			
		ИначеЕсли ВыбранноеЗначение.Свойство("СоответствиеОбъектовРемонта") Тогда
			
			Для каждого КлючИЗначение из ВыбранноеЗначение.СоответствиеОбъектовРемонта Цикл
				
				ОбъектРемонта = КлючИЗначение.Ключ;
				
				МассивСтрок = ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта", ОбъектРемонта));
				Если МассивСтрок.Количество() = 0 Тогда
					НовСтрокаОРР = ОбъектыРемонта.Добавить();
					НовСтрокаОРР.ОбъектРемонта = ОбъектРемонта;
					Элементы.ОбъектыРемонта.ТекущаяСтрока = НовСтрокаОРР.ПолучитьИдентификатор();
				КонецЕсли;
				
				Для каждого Показатель Из КлючИЗначение.Значение Цикл
					НайденныеСтроки = Объект.Показатели.НайтиСтроки(Новый Структура("ОбъектРемонта, Показатель", ОбъектРемонта, Показатель));
					Если НайденныеСтроки.Количество() = 0 Тогда
						СтрокаТЧ = Объект.Показатели.Добавить();
						СтрокаТЧ.Показатель = Показатель;
						СтрокаТЧ.ОбъектРемонта = ОбъектРемонта;
					КонецЕсли;
				КонецЦикла;
				
			КонецЦикла;
			
			ЗаполнитьГраничныеЗначенияПоказателей();
			ЗаполнитьПараметрыОформленияЗначенийКПОР();
			
		КонецЕсли;
		
	Иначе	
		СтруктураВозврата = ПолучитьЗначенияИзмеряемыхПоказателейНаСервере(ВыбранноеЗначение);
		Если СтруктураВозврата = Неопределено Тогда
			Возврат;
		КонецЕсли;
		
		// Проверка ОР на предмет уже выбранных 
		МассивСтрок = ОбъектыРемонта.НайтиСтроки(Новый Структура("ОбъектРемонта", ВыбранноеЗначение));
		Если МассивСтрок.Количество() = 0 Тогда
			НовСтрокаОРР = ОбъектыРемонта.Добавить();
			НовСтрокаОРР.ОбъектРемонта = ВыбранноеЗначение;
			Элементы.ОбъектыРемонта.ТекущаяСтрока = НовСтрокаОРР.ПолучитьИдентификатор();
						
			Для каждого Элем Из СтруктураВозврата Цикл
				СтрокаТЧ = Объект.Показатели.Добавить();
				ЗаполнитьЗначенияСвойств(СтрокаТЧ, Элем);
				СтрокаТЧ.ОбъектРемонта = НовСтрокаОРР.ОбъектРемонта;
			КонецЦикла;
			
		Иначе
			ШаблонСообщения = НСтр("ru = 'Оборудование ""%1"" уже есть в списке оборудования!'");
			ТекстСообщения = СтрШаблон(ШаблонСообщения, ВыбранноеЗначение);
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения);
		КонецЕсли;
		
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьКонтролируемыеПоказателиНаСервере(ТекущийОР)
	
	СтруктураВозврата = ПолучитьЗначенияИзмеряемыхПоказателейНаСервере(ТекущийОР);
	Если СтруктураВозврата = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Для каждого Элем Из СтруктураВозврата Цикл
		Если Объект.Показатели.НайтиСтроки(Новый Структура("ОбъектРемонта, Показатель", ТекущийОР, Элем.Показатель)).Количество() = 0  Тогда
			СтрокаТЧ = Объект.Показатели.Добавить();
			ЗаполнитьЗначенияСвойств(СтрокаТЧ, Элем);
			СтрокаТЧ.ОбъектРемонта = ТекущийОР;
		КонецЕсли; 	
	КонецЦикла;
	
КонецПроцедуры

&НаСервереБезКонтекста
Функция ПолучитьСписокПоказателейОбъектаРемонта(ОбъектРемонта)
	
	СписокОтбора = Новый СписокЗначений;
	
	МассивПоказателей = ПланыВидовХарактеристик.торо_ИзмеряемыеПоказателиОбъектовРемонта.ПолучитьСтруктуруИзмеряемыхПоказателейОбъектовРемонта(ОбъектРемонта);
	
	Для Каждого СтруктураПоказателя Из МассивПоказателей Цикл
		СписокОтбора.Добавить(СтруктураПоказателя.Показатель);
	КонецЦикла;
	
	Возврат СписокОтбора;
	
КонецФункции

&НаСервереБезКонтекста
Функция ПолучитьЗначенияИзмеряемыхПоказателейНаСервере(ОбъектРемонта, Показатель = Неопределено)
	
	МассивПоказателей = ПланыВидовХарактеристик.торо_ИзмеряемыеПоказателиОбъектовРемонта.ПолучитьСтруктуруИзмеряемыхПоказателейОбъектовРемонта(ОбъектРемонта,Показатель,Истина);
	
	Если ЗначениеЗаполнено(Показатель) Тогда
		Если МассивПоказателей.Количество() > 0 Тогда
			СтруктураВозврата = МассивПоказателей[0];
		Иначе
			СтруктураВозврата = Новый Структура;
		КонецЕсли;
	Иначе
		СтруктураВозврата = МассивПоказателей;
	КонецЕсли;
	
	Возврат СтруктураВозврата;
	
КонецФункции

&НаКлиенте
Процедура ПодборОРОбработчикОжидания()
	
	Если Объект.Показатели.Количество() = 0 ИЛИ ТипЗнч(Объект.ДокументОснование) = Тип("ДокументСсылка.торо_АктОВыполненииЭтапаРабот")
		ИЛИ ТипЗнч(Объект.ДокументОснование) = Тип("ДокументСсылка.торо_АктОВыполненииРегламентногоМероприятия") Тогда
		ОткрытьФормуПодбораОРИзОснования(Элементы.ОбъектыРемонта);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОткрытьФормуПодбораОРИзОснования(Элемент, ОдиночныйВыбор = Ложь)
	
	стрПараметров = Новый структура("ТипВыходногоДок, ДокументыОснования, ОдиночныйВыбор, ОткрыватьБезусловно, КонтролируемыеПоказатели", 
			ПредопределенноеЗначение("Перечисление.торо_ВидыВыходныхДокументов.УчетКонтролируемыхПоказателей"), Объект.ДокументОснование, ОдиночныйВыбор, Истина, Истина);		
			
	ОткрытьФорму("ОбщаяФорма.торо_ФормаПодбораОРПриВводеНаОсновании", стрПараметров, Элемент, ЭтаФорма.УникальныйИдентификатор, ВариантОткрытияОкна.ОтдельноеОкно);

КонецПроцедуры

#КонецОбласти