////////////////////////////////////////////////////////////////////////////////
// СЛУЖЕБНЫЕ ПЕРЕМЕННЫЕ
&НаКлиенте
Перем СотрудникУдаляемойСтроки;

#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	УстановитьПривилегированныйРежим(Истина);
	
	Если Параметры.Ключ.Пустая() Тогда
		
		// Очистим реквизит Сотрудник, если он уже принят
		Если ЗначениеЗаполнено(Параметры.ЗначениеКопирования) И Параметры.ЗначениеКопирования.Проведен Тогда
			Объект.Сотрудник = Справочники.Сотрудники.ПустаяСсылка();
			Объект.ФизическоеЛицо = Справочники.ФизическиеЛица.ПустаяСсылка();
		КонецЕсли; 
		
		Если Параметры.Свойство("Организация") И ЗначениеЗаполнено(Параметры.Организация) Тогда 
			Объект.Организация = Параметры.Организация;
		КонецЕсли;
		Если Параметры.Свойство("Сотрудник") И ЗначениеЗаполнено(Параметры.Сотрудник) Тогда 
			Объект.Сотрудник = Параметры.Сотрудник;
		КонецЕсли;
		Если Параметры.Свойство("СтатьяТКРФ") И ЗначениеЗаполнено(Параметры.СтатьяТКРФ) Тогда 
			Объект.СтатьяТКРФ = Параметры.СтатьяТКРФ;
		КонецЕсли;
		Если Параметры.Свойство("ДатаУвольнения") И ЗначениеЗаполнено(Параметры.ДатаУвольнения) Тогда 
			Объект.ДатаУвольнения = Параметры.ДатаУвольнения;
		ИначеЕсли ЗначениеЗаполнено(Параметры.ЗначениеКопирования) Тогда
			Объект.ДатаУвольнения = Параметры.ЗначениеКопирования.ДатаУвольнения;
		Иначе
			Объект.ДатаУвольнения = ТекущаяДатаСеанса();
		КонецЕсли;
		
		// заполнение нового документа
		ЗначенияДляЗаполнения = Новый Структура("Организация, Ответственный",
		"Объект.Организация",
		"Объект.Ответственный");
		ЗарплатаКадры.ЗаполнитьПервоначальныеЗначенияВФорме(ЭтаФорма, ЗначенияДляЗаполнения);
		
		ДатаУвольненияПриИзмененииНаСервере();
		
		УстановитьФункциональныеОпцииФормы();
		
	КонецЕсли;
	
	// СтандартныеПодсистемы.Свойства
	ДополнительныеПараметры = Новый Структура;
	ДополнительныеПараметры.Вставить("ИмяЭлементаДляРазмещения", "ГруппаДополнительныеРеквизиты");
	УправлениеСвойствами.ПриСозданииНаСервере(ЭтотОбъект, ДополнительныеПараметры);
	// Конец СтандартныеПодсистемы.Свойства

	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКоманды.ПриСозданииНаСервере(ЭтаФорма);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	// Заголовок формы++
	торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--

КонецПроцедуры

&НаСервере
Процедура ПриЧтенииНаСервере(ТекущийОбъект)
	
	// СтандартныеПодсистемы.Свойства
	УправлениеСвойствами.ПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.Свойства
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиентСервер.ОбновитьКоманды(ЭтотОбъект, Объект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	// СтандартныеПодсистемы.ДатыЗапретаИзменения
	ДатыЗапретаИзменения.ОбъектПриЧтенииНаСервере(ЭтотОбъект, ТекущийОбъект);
	// Конец СтандартныеПодсистемы.ДатыЗапретаИзменения
	
	УстановитьПривилегированныйРежим(Истина);
	
	ВыполнитьЧтениеНаСервере(ТекущийОбъект);
	
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
	
	// Заголовок формы++
		торо_РаботаСДиалогами.УстановитьЗаголовокФормыДокумента(Объект, ЭтаФорма, "");
	// Заголовок формы--
	
КонецПроцедуры

&НаКлиенте
Процедура ПослеЗаписи(ПараметрыЗаписи)
	
	ПараметрыОповещения = Новый Структура;
	ПараметрыОповещения.Вставить("Проведен", ЭтаФорма.Объект.Проведен);
	ПараметрыОповещения.Вставить("ПомеченНаУдаление", ЭтаФорма.Объект.ПометкаУдаления);
	ПараметрыОповещения.Вставить("Результат", ЭтаФорма.Объект.Ссылка);
	ПараметрыОповещения.Вставить("Ответственный", ЭтаФорма.Объект.Ответственный);
	ПараметрыОповещения.Вставить("ДатаДокумента", ЭтаФорма.Объект.Дата);
	ПараметрыОповещения.Вставить("НомерДокумента", ЭтаФорма.Объект.Номер);
	ПараметрыОповещения.Вставить("Сотрудник", ЭтаФорма.Объект.Сотрудник);
	Оповестить("ПослеЗаписиОбъектаСвязанногоСПереводомКДругомуРаботодателю", ПараметрыОповещения, ЭтаФорма);
	
	КадровыйУчетКлиент.ОповеститьОбИзмененииРабочегоМеста(ЭтаФорма);
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
Процедура ДатаУвольненияПриИзменении(Элемент)
	ДатаУвольненияПриИзмененииНаСервере()	
КонецПроцедуры

&НаКлиенте
Процедура КомментарийНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ОбщегоНазначенияКлиент.ПоказатьФормуРедактированияКомментария(Элемент.ТекстРедактирования, ЭтотОбъект, "Объект.Комментарий");	
КонецПроцедуры

&НаКлиенте
Процедура РуководительПриИзменении(Элемент)
		
	Если ЗначениеЗаполнено(Объект.Руководитель) Тогда 
		Объект.ДолжностьРуководителя = ПолучитьДолжностьРуководителя(Объект.Руководитель, Объект.Дата);
	Иначе 
		Объект.ДолжностьРуководителя = Неопределено;
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура СотрудникАвтоПодбор(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, Ожидание, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ДанныеВыбора = Новый СписокЗначений;
	
	масСотрудников = СписокПодходящихСорудников(Объект.ДатаУвольнения, Объект.ДатаУвольнения, Объект.Организация, Текст);
	Для каждого ТекСтрока из масСотрудников Цикл
		ДанныеВыбора.Добавить(ТекСтрока);
	КонецЦикла;
КонецПроцедуры

&НаКлиенте
Процедура СотрудникОкончаниеВводаТекста(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	масСотрудников = СписокПодходящихСорудников(Объект.ДатаУвольнения, Объект.ДатаУвольнения, Объект.Организация, Текст);
	ДанныеВыбора = Новый СписокЗначений;
	Для каждого ТекСтрока из масСотрудников Цикл
		ДанныеВыбора.Добавить(ТекСтрока);
	КонецЦикла;
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

&НаСервереБезКонтекста
Функция СписокПодходящихСорудников(Дата, Дата1, Организация, Текст)
	Возврат Документы.Увольнение.СписокПодходящихСорудников(Дата, Дата, Организация, Текст);
КонецФункции

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
Процедура ВыполнитьЧтениеНаСервере(ТекущийОбъект)
		
	УстановитьФункциональныеОпцииФормы();
	
КонецПроцедуры

&НаСервере
Процедура УстановитьФункциональныеОпцииФормы()
	
	ПараметрыФО = Новый Структура("Организация", Объект.Организация);
	УстановитьПараметрыФункциональныхОпцийФормы(ПараметрыФО);
	
КонецПроцедуры

&НаСервере
Процедура УстановитьПодразделение()
	КадровыеДанныеСотрудников = КадровыйУчет.КадровыеДанныеСотрудников(Истина, ОбщегоНазначенияКлиентСервер.ЗначениеВМассиве(Объект.Сотрудник), "ТекущееПодразделение", Объект.ДатаУвольнения);
	Если КадровыеДанныеСотрудников <> Неопределено и КадровыеДанныеСотрудников.Количество() > 0 Тогда
		Объект.Подразделение = КадровыеДанныеСотрудников[0]["ТекущееПодразделение"];
	КонецЕсли;
КонецПроцедуры

// Период начислений в строках документа

//////////////////////////////////////////////////////////////////////
// Обслуживание элементов формы

&НаСервере
Процедура ДатаУвольненияПриИзмененииНаСервере()
	
	ПриИзмененииСотрудникаИлиДатыУвольнения();	
	
КонецПроцедуры

&НаСервере
Процедура ПриИзмененииСотрудникаИлиДатыУвольнения()
	
	УстановитьПривилегированныйРежим(Истина);
	
	УстановитьПодразделение();
	
КонецПроцедуры

&НаКлиенте
Процедура СотрудникНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	
	СписокСотрудников = Новый СписокЗначений;
	масСотрудников = СписокПодходящихСорудников(Объект.ДатаУвольнения, Объект.ДатаУвольнения, Объект.Организация, "");
	Для каждого Сотрудник Из масСотрудников Цикл
		СписокСотрудников.Добавить(Сотрудник);
	КонецЦикла; 
	
	СтруктураОтбора = Новый Структура("СписокСотрудников", СписокСотрудников);
	ПараметрыФормы = Новый Структура("Отбор, РежимВыбора", СтруктураОтбора, Истина);
	ОткрытьФорму("Справочник.Сотрудники.ФормаВыбора", ПараметрыФормы, Элемент,,,,,РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
	
КонецПроцедуры

&НаСервере 
Функция ПолучитьДолжностьРуководителя(Руководитель, Дата)
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	КадроваяИсторияСотрудниковСрезПоследних.Должность КАК Должность
				   |ИЗ
	               |	РегистрСведений.КадроваяИсторияСотрудников.СрезПоследних(&Дата, ФизическоеЛицо = &Руководитель) КАК КадроваяИсторияСотрудниковСрезПоследних";
	
	Запрос.УстановитьПараметр("Дата", Дата);
	Запрос.УстановитьПараметр("Руководитель", Руководитель);
	Выборка = Запрос.Выполнить().Выбрать();
	
	Если Выборка.Следующий() Тогда
		Должность = Выборка.Должность;
	КонецЕсли;
	
	Возврат Должность;
		
КонецФункции

#КонецОбласти