#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Если Не Параметры.Свойство("Ключ") Или Не ЗначениеЗаполнено(Параметры.Ключ) Тогда
		
		Параметры.Свойство("АдресСпискаПодобранныхСотрудников", АдресСпискаПодобранныхСотрудников);
		
		РедактированиеСтрокиСписочногоДокумента = Истина;
		
		УстановитьПривилегированныйРежим(Истина);
		ДокументОбъект = Документы.КадровыйПеревод.СоздатьДокумент();
		УстановитьПривилегированныйРежим(Ложь);
		
		ДокументОбъект.Заполнить(Неопределено);
		ЗначениеВРеквизитФормы(ДокументОбъект, "Объект");
		
		СтрокаСписочногоДокумента = Параметры.СтрокаСписочногоДокумента;
		ЗарплатаКадрыРасширенныйКлиентСервер.ЗаполнитьЗначенияЗаполненныхСвойств(Объект, СтрокаСписочногоДокумента);
		
		СсылкаНаОбъект = СтрокаСписочногоДокумента.СсылкаНаОбъект;
		
		ПоложениеКоманднойПанели = ПоложениеКоманднойПанелиФормы.Нет; 
		АвтоЗаголовок = Ложь;
		
		Заголовок = НСтр("ru = 'Кадровый перевод';
						|en = 'Employee transfer'");
		
		ЭлементВидЗанятости = Элементы.Найти("ВидЗанятости");
		Если ЭлементВидЗанятости <> Неопределено Тогда
			
			ОбновитьСвязиПараметровВыбораЭлемента = Ложь;
			СвязиПараметровВыбораЭлемента = Новый Массив(ЭлементВидЗанятости.СвязиПараметровВыбора);
			Для ИндексЭлемента = 0 По СвязиПараметровВыбораЭлемента.Количество() - 1 Цикл
				
				СвязьПараметров = СвязиПараметровВыбораЭлемента.Получить(ИндексЭлемента);
				Если СвязьПараметров.Имя = "Отбор.ТекущийВидЗанятости" Тогда
					
					ОбновитьСвязиПараметровВыбораЭлемента = Истина;
					СвязиПараметровВыбораЭлемента.Удалить(ИндексЭлемента);
					Прервать;
					
				КонецЕсли;
				
			КонецЦикла;
			
			Если ОбновитьСвязиПараметровВыбораЭлемента Тогда
				ЭлементВидЗанятости.СвязиПараметровВыбора = Новый ФиксированныйМассив(СвязиПараметровВыбораЭлемента);
			КонецЕсли;
			
		КонецЕсли;
		
		Если Не ЗначениеЗаполнено(Объект.ДатаНачала) Тогда
			Объект.ДатаНачала = ТекущаяДатаСеанса();
		КонецЕсли;
		
		Если Параметры.Свойство("Сотрудник") И ЗначениеЗаполнено(Параметры.Сотрудник) Тогда 
			Объект.Сотрудник = Параметры.Сотрудник;
		КонецЕсли;
		
		Если ЗначениеЗаполнено(Объект.Сотрудник) Тогда
			УстановитьТекущиеДанныеСотрудника();
		КонецЕсли;
		
	КонецЕсли;
	
	УстановитьДоступностьЭлементов(ЭтаФорма);
	
КонецПроцедуры

&НаСервере
Процедура ОбработкаПроверкиЗаполненияНаСервере(Отказ, ПроверяемыеРеквизиты)
	
	УстановитьПривилегированныйРежим(Истина);
	
	ТекущийОбъект = РеквизитФормыВЗначение("Объект");
	
	УстановитьПривилегированныйРежим(Ложь);
	
	Если РедактированиеСтрокиСписочногоДокумента Тогда
		ТекущийОбъект.ДополнительныеСвойства.Вставить("ПроверкаСтрокиСписочногоДокумента", Истина);
		ПараметрыСеанса.ИнтерактивнаяПроверкаЗаполненияСвойств = Истина;
	КонецЕсли;
	
	Если Не ТекущийОбъект.ПроверитьЗаполнение() Тогда
		Отказ = Истина;
	КонецЕсли;
	
	ОбщегоНазначенияКлиентСервер.УдалитьЗначениеИзМассива(ПроверяемыеРеквизиты, "Объект");
	
КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытием(Отказ, ЗавершениеРаботы, ТекстПредупреждения, СтандартнаяОбработка)
	
	Оповещение = Новый ОписаниеОповещения("ПередЗакрытиемЗавершение", ЭтотОбъект);
	ОбщегоНазначенияКлиент.ПоказатьПодтверждениеЗакрытияФормы(Оповещение, Отказ, ЗавершениеРаботы);
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура СотрудникПриИзменении(Элемент)
	
	СотрудникПриИзмененииНаСервере();
	
КонецПроцедуры

&НаКлиенте
Процедура СотрудникНачалоВыбора(Элемент, ДанныеВыбора, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	
	СписокСотрудников = Новый СписокЗначений;
	масСотрудников = СписокПодходящихСорудников(Объект.ДатаНачала, Объект.ДатаОкончания, Объект.Организация, "");
	Для каждого Сотрудник Из масСотрудников Цикл
		СписокСотрудников.Добавить(Сотрудник);
	КонецЦикла; 
	
	СтруктураОтбора = Новый Структура("СписокСотрудников", СписокСотрудников);
	ПараметрыФормы = Новый Структура("Отбор, РежимВыбора", СтруктураОтбора, Истина);
	ОткрытьФорму("Справочник.Сотрудники.ФормаВыбора", ПараметрыФормы, Элемент,,,,,РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
КонецПроцедуры

&НаКлиенте
Процедура СотрудникАвтоПодбор(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, Ожидание, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ДанныеВыбора = Новый СписокЗначений;
	
	масСотрудников = СписокПодходящихСорудников(Объект.ДатаНачала, Объект.ДатаОкончания, Объект.Организация, Текст);
	Для каждого ТекСтрока из масСотрудников Цикл
		ДанныеВыбора.Добавить(ТекСтрока);
	КонецЦикла;
КонецПроцедуры

&НаКлиенте
Процедура СотрудникОкончаниеВводаТекста(Элемент, Текст, ДанныеВыбора, ПараметрыПолученияДанных, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ДанныеВыбора = Новый СписокЗначений;
	
	масСотрудников = СписокПодходящихСорудников(Объект.ДатаНачала, Объект.ДатаОкончания, Объект.Организация, Текст);
	Для каждого ТекСтрока из масСотрудников Цикл
		ДанныеВыбора.Добавить(ТекСтрока);
	КонецЦикла;
КонецПроцедуры

&НаКлиенте
Процедура ДатаНачалаПриИзменении(Элемент)
	
	ДатаНачалаПриИзмененииНаСервере();
	
КонецПроцедуры

&НаКлиенте
Процедура ИзмениласьПозицияПриИзменении(Элемент)
	
	ИзмениласьПозицияПриИзмененииНаСервере();
	
КонецПроцедуры

&НаКлиенте
Процедура ПодразделениеПриИзменении(Элемент)
	
	ПодразделениеПриИзмененииНаСервере();
	
КонецПроцедуры

&НаКлиенте
Процедура ИзменилсяГрафикПриИзменении(Элемент)
	
	ЗаполнитьГрафик();
	
КонецПроцедуры

&НаКлиенте
Процедура ИзменитьСведенияОДоговореПриИзменении(Элемент)
	Объект.ДатаЗавершенияТрудовогоДоговора = ТекущаяДатаЗавершенияТД;
	УстановитьДоступностьЭлементов(ЭтаФорма);
	Элементы.ДатаЗавершенияТДПодсказка.Видимость = Объект.ИзменитьСведенияОДоговоре;
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ОК(Команда)
	
	ЗавершитьРедактированиеСтрокиДокумента();
	
КонецПроцедуры

&НаКлиенте
Процедура ЗавершитьРедактированиеСтрокиДокумента()
	
	ОчиститьСообщения();
	
	Отказ = Ложь;
	Если Модифицированность Тогда
		
		Если Не ПроверитьЗаполнение() Тогда
			Отказ = Истина;
		Иначе
			ОписаниеДокумента = СтруктураСОписаниемДокумента();
			Модифицированность = Ложь;
		КонецЕсли;
		
	Иначе
		ОписаниеДокумента = Неопределено;
	КонецЕсли;
	
	Если Не Отказ Тогда
		Закрыть(ОписаниеДокумента);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти


#Область СлужебныеПроцедурыИФункции

&НаКлиентеНаСервереБезКонтекста
Процедура УстановитьДоступностьЭлементов(Форма)
	
	ИзменитьПозициюДолжность		= Форма.Объект.ИзменитьПодразделениеИДолжность;
	ИзменитьГрафикРаботы			= Форма.Объект.ИзменитьГрафикРаботы;
	ИзменитьСведенияОДоговоре		= Форма.Объект.ИзменитьСведенияОДоговоре;
	
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"Подразделение",
		"Доступность",
		ИзменитьПозициюДолжность);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
			Форма.Элементы,
			"Должность",
			"Доступность",
			ИзменитьПозициюДолжность);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"КоличествоСтавок",
		"Доступность",
		ИзменитьПозициюДолжность);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"РазрядКатегория",
		"Доступность",
		ИзменитьПозициюДолжность);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"ВидЗанятости",
		"Доступность",
		ИзменитьПозициюДолжность);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"СтатьяФинансирования",
		"Доступность",
		ИзменитьПозициюДолжность);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"График",
		"Доступность",
		ИзменитьГрафикРаботы);
		
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"ДатаЗавершенияТрудовогоДоговора",
		"Доступность",
		ИзменитьСведенияОДоговоре);
		
КонецПроцедуры

&НаСервере
Процедура УстановитьДанныеРабочегоМеста()
	
	УстановитьПривилегированныйРежим(Истина);
	
	Если ЗначениеЗаполнено(Объект.Сотрудник) Тогда
		Объект.Подразделение 				= ТекущееПодразделение;
		Объект.Должность 					= ТекущаяДолжность;
		Объект.ВидЗанятости 				= ТекущийВидЗанятости;
	Иначе
		Если ЗначениеЗаполнено(Объект.Сотрудник) И НЕ ЗначениеЗаполнено(Объект.ОбособленноеПодразделение) Тогда
			Объект.ОбособленноеПодразделение 	= Объект.Организация;
		КонецЕсли;
		Объект.Подразделение 				= Справочники.СтруктураПредприятия.ПустаяСсылка();
		Объект.Должность 					= Справочники.Должности.ПустаяСсылка();
		Объект.ВидЗанятости 				= Перечисления.ВидыЗанятости.ПустаяСсылка();
	КонецЕсли;
	
	УстановитьПривилегированныйРежим(Ложь);
	
КонецПроцедуры

&НаСервере
Процедура ИзмениласьПозицияПриИзмененииНаСервере()
	
	УстановитьДанныеРабочегоМеста();
	УстановитьДоступностьЭлементов(ЭтаФорма);
	УстановитьКомментарии(ЭтаФорма);
	
КонецПроцедуры

&НаСервере
Процедура СотрудникПриИзмененииНаСервере()
	
	УстановитьТекущиеДанныеСотрудника();
	
КонецПроцедуры

&НаСервере
Процедура УстановитьТекущиеДанныеСотрудника()
	
	УстановитьТекущиеКадровыеДанные();
	УстановитьДанныеРабочегоМеста();
	
	Объект.Подразделение = ТекущееПодразделение;
	Объект.Должность = ТекущаяДолжность;
	Объект.КоличествоСтавок = ТекущееКоличествоСтавок;
	Объект.ГрафикРаботы = ТекущийГрафикРаботы;
	Объект.ВидЗанятости = ТекущийВидЗанятости;
	Объект.ФизическоеЛицо = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(Объект.Сотрудник, "ФизическоеЛицо");
	Объект.ДатаЗавершенияТрудовогоДоговора = ПолучитьДатуЗавершенияТД(Объект.Сотрудник, Объект.ДатаНачала);
	
	УстановитьДоступностьЭлементовИзмененияДанныхСотрудника();
	
	УстановитьДоступностьНовогоПодразделения(ЭтаФорма);
	
КонецПроцедуры

&НаСервере
Процедура УстановитьДоступностьЭлементовИзмененияДанныхСотрудника()
	
	ДоступноИзменениеДанныхСотрудника = ЗначениеЗаполнено(Объект.Сотрудник);
	
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Элементы,
		"ИзмениласьПозиция",
		"Доступность",
		ДоступноИзменениеДанныхСотрудника);
	
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Элементы,
		"ИзменитьВидЗанятости",
		"Доступность",
		ДоступноИзменениеДанныхСотрудника);
	
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Элементы,
		"ИзменилсяГрафик",
		"Доступность",
		ДоступноИзменениеДанныхСотрудника);
	
КонецПроцедуры

&НаКлиентеНаСервереБезКонтекста
Процедура УстановитьДоступностьНовогоПодразделения(Форма)
	
	ДоступностьПодразделнияИДолжности = 
		Форма.Объект.ИзменитьПодразделениеИДолжность;
	
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"Подразделение",
		"Доступность",
		ДоступностьПодразделнияИДолжности);
	
	ОбщегоНазначенияКлиентСервер.УстановитьСвойствоЭлементаФормы(
		Форма.Элементы,
		"Должность",
		"Доступность",
		ДоступностьПодразделнияИДолжности);
	
КонецПроцедуры

&НаСервере
Процедура ДатаНачалаПриИзмененииНаСервере()
	
	УстановитьТекущиеДанныеСотрудника();
	
КонецПроцедуры

&НаСервере
Процедура УстановитьТекущиеКадровыеДанные()
	
	ПериодПредыдущихЗначений = '00010101';
	Если ЗначениеЗаполнено(Объект.Сотрудник) Тогда
		
		СтруктураОтбораКадровыхДанных = Новый Структура("ЛевоеЗначение, ВидСравнения, ПравоеЗначение", "Регистратор", "НЕ В", СсылкаНаОбъект);
		Отбор = Новый Массив;
		Отбор.Добавить(СтруктураОтбораКадровыхДанных);
		
		ПоляОтбораПериодическихДанных = Новый Структура;
		ПоляОтбораПериодическихДанных.Вставить("КадроваяИсторияСотрудников", Отбор);
		ПоляОтбораПериодическихДанных.Вставить("ГрафикРаботыСотрудников", Отбор);
		ПоляОтбораПериодическихДанных.Вставить("РазрядыКатегорииСотрудников", Отбор);
		ПоляОтбораПериодическихДанных.Вставить("ТерриторииСотрудников", Отбор);
		ПоляОтбораПериодическихДанных.Вставить("ВидыЗанятостиСотрудников", Отбор);
		Поля = "Подразделение,Должность,КоличествоСтавок,ГрафикРаботы,Организация,ВидЗанятости";
	 
		УстановитьПривилегированныйРежим(Истина);
		
		ДанныеСотрудников = КадровыйУчет.КадровыеДанныеСотрудников(Ложь, Объект.Сотрудник, Поля, ВремяРегистрации, ПоляОтбораПериодическихДанных);
		
		УстановитьПривилегированныйРежим(Ложь);
		
		Если ДанныеСотрудников.Количество() > 0 Тогда
			
			ТекущиеКадровыеДанныеСотрудника = ДанныеСотрудников[0];
			
			ТекущееПодразделение 				= ТекущиеКадровыеДанныеСотрудника.Подразделение;
			ТекущийГрафикРаботы 				= ТекущиеКадровыеДанныеСотрудника.ГрафикРаботы;
			ТекущееКоличествоСтавок 			= ТекущиеКадровыеДанныеСотрудника.КоличествоСтавок;
			ТекущаяДолжность 						= ТекущиеКадровыеДанныеСотрудника.Должность;
			ТекущийВидЗанятости 					= ТекущиеКадровыеДанныеСотрудника.ВидЗанятости;
			ТекущаяДатаЗавершенияТД = ПолучитьДатуЗавершенияТД(Объект.Сотрудник, Объект.ДатаНачала);
			
		КонецЕсли;
		
	Иначе
		
		ТекущееПодразделение 					= Справочники.СтруктураПредприятия.ПустаяСсылка();
		ТекущаяДолжность 						= Справочники.Должности.ПустаяСсылка();
		ТекущийГрафикРаботы 					= Справочники.Календари.ПустаяСсылка();
		ТекущееКоличествоСтавок 				= 0;
		ТекущийВидЗанятости 					= Перечисления.ВидыЗанятости.ПустаяСсылка();
		ТекущаяДатаЗавершенияТД = Неопределено;
		
	КонецЕсли;
	
	УстановитьКомментарии(ЭтаФорма);
	
КонецПроцедуры

&НаКлиентеНаСервереБезКонтекста
Процедура УстановитьКомментарии(Форма)
	
	Форма.КоличествоСтавокКомментарий = КадровыйУчетРасширенныйКлиентСервер.КомментарийККоличествуСтавок(
		Форма.ТекущееКоличествоСтавок,
		Форма.Объект.КоличествоСтавок,
		Форма.Объект.ИзменитьПодразделениеИДолжность);
	
	УстановитьПозицияКомментарий(Форма);
	
	УстановитьГрафикРаботыКомментарий(Форма);
	
	Форма.Элементы.ПозицияКомментарий.Подсказка = Форма.ПозицияКомментарий;
	Форма.Элементы.ГрафикРаботыКомментарий.Подсказка = Форма.ГрафикРаботыКомментарий;
	
	Форма.Элементы.ДатаЗавершенияТДПодсказка.Видимость = Форма.Объект.ИзменитьСведенияОДоговоре;
	
КонецПроцедуры

&НаСервереБезКонтекста
Процедура УстановитьПозицияКомментарий(Форма)
	
	Если Форма.Объект.ИзменитьПодразделениеИДолжность Тогда
		
		Если ЗначениеЗаполнено(Форма.ТекущееПодразделение) Тогда
			
			Отбор = Новый Структура("Сотрудник", Форма.Объект.Сотрудник);
			ТабПредыдущихДанных = РегистрыСведений.КадроваяИсторияСотрудников.СрезПоследних(Форма.Объект.ДатаНачала-1, Отбор);
			Если ТабПредыдущихДанных.Количество() > 0 Тогда 
				ПредыдущиеДанные = ТабПредыдущихДанных[0];
				Если ЗначениеЗаполнено(ПредыдущиеДанные.ДействуетДо) И Форма.Объект.ДатаНачала > ПредыдущиеДанные.ДействуетДо 
				  И ЗначениеЗаполнено(ПредыдущиеДанные.ДолжностьПоОкончании) Тогда
					ПредыдущаяДолжность = ПредыдущиеДанные.ДолжностьПоОкончании;
					ПредыдущееПодразделение = ПредыдущиеДанные.ПодразделениеПоОкончании;
				Иначе
					ПредыдущаяДолжность = ПредыдущиеДанные.Должность;
					ПредыдущееПодразделение = ПредыдущиеДанные.Подразделение; 
				КонецЕсли;
				
				ПозицияКомментарий = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
				НСтр("ru = 'Ранее сотрудник занимал должность ""%1"" в подразделении ""%2""'"),
				ПредыдущаяДолжность, ПредыдущееПодразделение);
			КонецЕсли;
					
		Иначе
			
			Форма.ПозицияКомментарий = НСтр("ru = 'Сотрудник еще не принят на работу';
											|en = 'Employee has not been recruited yet'");
			
		КонецЕсли;
		
	Иначе
		
		Форма.ПозицияКомментарий = "";
		
	КонецЕсли;	
	
КонецПроцедуры

&НаКлиентеНаСервереБезКонтекста
Процедура УстановитьГрафикРаботыКомментарий(Форма)
	
	Если Форма.Объект.ИзменитьГрафикРаботы Тогда
			
		Форма.ГрафикРаботыКомментарий = СтроковыеФункцииКлиентСервер.ПодставитьПараметрыВСтроку(
			НСтр("ru = 'Ранее сотрудник работал по графику %1';
				|en = 'Earlier the employee worked according to schedule %1'"),
			Форма.ТекущийГрафикРаботы);

	Иначе
				
		Форма.ГрафикРаботыКомментарий = "";
		
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура ПодразделениеПриИзмененииНаСервере()
	
	ЗаполнитьГрафик(Истина);
	
КонецПроцедуры

&НаСервере
Процедура ЗаполнитьГрафик(Перезаполнение = Ложь)
	
	Если Объект.ИзменитьГрафикРаботы ИЛИ Перезаполнение Тогда
		
		ЗначенияДляЗаполнения = Новый Структура;
		
		ЗначенияДляЗаполнения.Вставить("Подразделение", "Объект.Подразделение");
		ЗначенияДляЗаполнения.Вставить("Организация",	"Объект.Организация");
		ЗначенияДляЗаполнения.Вставить("ГрафикРаботы",	"Объект.ГрафикРаботы");
		
		ФиксированныеЗначения = Новый Массив;
		ФиксированныеЗначения.Добавить("Организация");
		ФиксированныеЗначения.Добавить("Подразделение");
		
		ЗарплатаКадры.ЗаполнитьЗначенияВФорме(ЭтаФорма, ЗначенияДляЗаполнения, ФиксированныеЗначения);
		
		Если Не ЗначениеЗаполнено(Объект.ГрафикРаботы) Тогда
			Объект.ГрафикРаботы = ТекущийГрафикРаботы;
		КонецЕсли;
		
		Если Перезаполнение Тогда
			
			Если Объект.ГрафикРаботы <> ТекущийГрафикРаботы Тогда
				Объект.ИзменитьГрафикРаботы = Истина;
			Иначе
				Объект.ИзменитьГрафикРаботы = Ложь;
			КонецЕсли;
			
		КонецЕсли;
		
	Иначе
		
		Объект.ГрафикРаботы = ТекущийГрафикРаботы;
		
	КонецЕсли;
	
	УстановитьДоступностьЭлементов(ЭтаФорма);
	УстановитьКомментарии(ЭтаФорма);
	
КонецПроцедуры

#Область РедактированиеДанныхКонтрактаДоговора

&НаСервере
Функция СтруктураСОписаниемДокумента()
	ОписаниеДокумента = ЗарплатаКадрыРасширенныйВызовСервера.СтруктураПоМетаданным("Документ.КадровыйПеревод");
	ЗарплатаКадрыРасширенный.ЗаполнитьОбъектПоОбразцу(ОписаниеДокумента, Объект);
	
	Возврат ОписаниеДокумента;
	
КонецФункции

&НаСервереБезКонтекста
Функция СписокПодходящихСорудников(ДатаНачала, ДатаОкончания, Организация, Текст)
	Возврат Документы.КадровыйПеревод.СписокПодходящихСорудников(ДатаНачала, ДатаОкончания, Организация, Текст);
КонецФункции

&НаСервереБезКонтекста
Функция ПолучитьДатуЗавершенияТД(Сотрудник, ДатаНачала)
	Возврат Документы.КадровыйПеревод.ПолучитьДатуЗавершенияТД(Сотрудник, ДатаНачала);
КонецФункции

#КонецОбласти

&НаКлиенте
Процедура ПередЗакрытиемЗавершение(Результат, ДополнительныеПараметры) Экспорт 
	ЗавершитьРедактированиеСтрокиДокумента();
КонецПроцедуры

#КонецОбласти
