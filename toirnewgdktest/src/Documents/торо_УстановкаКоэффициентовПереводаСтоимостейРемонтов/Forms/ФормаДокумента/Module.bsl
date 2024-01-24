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
	
	Если Параметры.Свойство("ВидОперации") Тогда
		Объект.ВидОперации = Параметры.ВидОперации;
	Иначе
		Если НЕ ЗначениеЗаполнено(Объект.Ссылка) Тогда
			Объект.ВидОперации = Перечисления.торо_ВидыДокументаВводНачДанных.ПоОбъектуРемонта;
		КонецЕсли;
	КонецЕсли;
	
	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, Строка(Объект.ВидОперации));
	// Заголовок формы--
	
	Если Объект.ВидОперации = Перечисления.торо_ВидыДокументаВводНачДанных.ПоОбъектуРемонта Тогда
		Элементы.СтраницыПоВидуОперации.ПодчиненныеЭлементы.СтраницаСпискиОбъектов.Видимость = Ложь;
		Элементы.СтраницыПоВидуОперации.ПодчиненныеЭлементы.СтраницаОбъектыРемонта.Видимость = Истина;
	Иначе
		Элементы.СтраницыПоВидуОперации.ПодчиненныеЭлементы.СтраницаСпискиОбъектов.Видимость = Истина;
		Элементы.СтраницыПоВидуОперации.ПодчиненныеЭлементы.СтраницаОбъектыРемонта.Видимость = Ложь;
	КонецЕсли;
	
	Если НЕ ЗначениеЗаполнено(Объект.Ссылка) Тогда
		
		Если НЕ ЗначениеЗаполнено(Объект.Организация) Тогда
			
			Объект.Организация = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить(
			"НастройкиТОиР",
			"ОсновнаяОрганизация",
			Истина);
			
		КонецЕсли;
		
		Если НЕ ЗначениеЗаполнено(Объект.Ответственный) Тогда
			
			Объект.Ответственный = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить(
			"НастройкиТОиР",
			"ОсновнойОтветственный",
			Справочники.Пользователи.ПустаяСсылка());
			
		КонецЕсли;
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
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, Строка(Объект.ВидОперации));
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

#Область ОбработчикиСобытийЭлементовТаблицыФормыКоэффициентыПеревода
&НаКлиенте
Процедура КоэффициентыПереводаПодразделениеНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	ТекСтрока = Элементы.КоэффициентыПеревода.ТекущиеДанные;
	
	СтандартнаяОбработка = ЗначениеЗаполнено(ТекСтрока.Организация);
	
	Если Не СтандартнаяОбработка Тогда
		ТекстСообщения = НСтр("ru = 'Нельзя выбирать подразделение без выбранной организации!'");
		ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыПереводаОбъектРемонтаОбработкаВыбора(Элемент, ВыбранноеЗначение, СтандартнаяОбработка)
	Если ВыбранноеЗначение <> Неопределено Тогда
		СтандартнаяОбработка = Ложь;
		
		ТекСтрока = Элементы.КоэффициентыПеревода.ТекущиеДанные;
		
		СтруктураОР = ПолучитьСтруктуруОР(ВыбранноеЗначение);
		
		ЕстьОшибка = Ложь;
		
		Если СтруктураОР.Организация <> ТекСтрока.Организация Тогда
			ЕстьОшибка = Истина;
		КонецЕсли;
		
		Если СтруктураОР.Подразделение <> ТекСтрока.Подразделение Тогда
			ЕстьОшибка = Истина;
		КонецЕсли;
		
		Если СтруктураОР.Направление <> ТекСтрока.Направления Тогда
			ЕстьОшибка = Истина;
		КонецЕсли;
		
		Если ЕстьОшибка Тогда
			Ответ = Неопределено;

			ПоказатьВопрос(Новый ОписаниеОповещения("КоэффициентыПереводаОбъектРемонтаОбработкаВыбораЗавершение", ЭтотОбъект, Новый Структура("ВыбранноеЗначение, СтруктураОР, ТекСтрока", ВыбранноеЗначение, СтруктураОР, ТекСтрока)), НСтр("ru = 'Заполнить значения реквизитов строки значениями из объекта ремонта?'"), РежимДиалогаВопрос.ДаНетОтмена);
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыПереводаОбъектРемонтаОбработкаВыбораЗавершение(РезультатВопроса, ДополнительныеПараметры) Экспорт
    
    ВыбранноеЗначение = ДополнительныеПараметры.ВыбранноеЗначение;
    СтруктураОР = ДополнительныеПараметры.СтруктураОР;
    ТекСтрока = ДополнительныеПараметры.ТекСтрока;
    
    
    Ответ = РезультатВопроса;
    Если Ответ = КодВозвратаДиалога.Да Тогда
        ЗаполнитьЗначенияСвойств(ТекСтрока,ВыбранноеЗначение);
        ТекСтрока.Организация = СтруктураОР.Организация;
        ТекСтрока.Подразделение = СтруктураОР.Подразделение;
        ТекСтрока.Направления = СтруктураОР.Направление;
        ТекСтрока.ОбъектРемонта = ВыбранноеЗначение;
    ИначеЕсли Ответ = КодВозвратаДиалога.Нет Тогда
        ТекСтрока.ОбъектРемонта = ВыбранноеЗначение;
    КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура КоэффициентыПереводаКлассификаторРемонтовНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	ТекСтрока = Элементы.КоэффициентыПеревода.ТекущиеДанные;
	
	СтандартнаяОбработка = ЗначениеЗаполнено(ТекСтрока.Направления) ИЛИ ЗначениеЗаполнено(ТекСтрока.Организация) ИЛИ ЗначениеЗаполнено(ТекСтрока.ОбъектРемонта);
	
	Если Не СтандартнаяОбработка Тогда
		ТекстСообщения = НСтр("ru = 'Нельзя выбирать вид обслуживания без выбранного направления или организации или объекта ремонта!'");
		ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
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
Функция ПолучитьСтруктуруОР(ОР)
	
	Возврат Новый Структура("Организация, Подразделение, Направление", ОР.Организация, ОР.Подразделение, ОР.Направление);
	
КонецФункции

&НаКлиенте
Процедура КомментарийНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	ОбщегоНазначенияКлиент.ПоказатьФормуРедактированияКомментария(
		Элемент.ТекстРедактирования, ЭтотОбъект, "Объект.Комментарий");	
	
КонецПроцедуры

#КонецОбласти