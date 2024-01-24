#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ПрограммныйИнтерфейс
// Подсистема "Управление доступом"

// Процедура ЗаполнитьНаборыЗначенийДоступа по свойствам объекта заполняет наборы значений доступа
// в таблице с полями:
//    НомерНабора     - Число                                     (необязательно, если набор один),
//    ВидДоступа      - ПланВидовХарактеристикСсылка.ВидыДоступа, (обязательно),
//    ЗначениеДоступа - Неопределено, СправочникСсылка или др.    (обязательно),
//    Чтение          - Булево (необязательно, если набор для всех прав) устанавливается для одной строки набора,
//    Добавление      - Булево (необязательно, если набор для всех прав) устанавливается для одной строки набора,
//    Изменение       - Булево (необязательно, если набор для всех прав) устанавливается для одной строки набора,
//    Удаление        - Булево (необязательно, если набор для всех прав) устанавливается для одной строки набора.
//
//  Вызывается из процедуры УправлениеДоступомСлужебный.ЗаписатьНаборыЗначенийДоступа(),
// если объект зарегистрирован в "ПодпискаНаСобытие.ЗаписатьНаборыЗначенийДоступа" и
// из таких же процедур объектов, у которых наборы значений доступа зависят от наборов этого
// объекта (в этом случае объект зарегистрирован в "ПодпискаНаСобытие.ЗаписатьЗависимыеНаборыЗначенийДоступа").
//
// Параметры:
//  Таблица - ТаблицаЗначений - ТабличнаяЧасть,
//                 РегистрСведенийНаборЗаписей.НаборыЗначенийДоступа,
//                 ТаблицаЗначений - возвращаемая УправлениеДоступом.ТаблицаНаборыЗначенийДоступа().
//
Процедура ЗаполнитьНаборыЗначенийДоступа(Таблица) Экспорт
	
	ЗарплатаКадры.ЗаполнитьНаборыПоОрганизацииИФизическимЛицам(ЭтотОбъект, Таблица, "Организация", "ФизическоеЛицо");
	
КонецПроцедуры
#КонецОбласти

#Область ОбработчикиСобытий

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	Документы.КадровыйПеревод.ОбработкаПроверкиЗаполнения(ЭтотОбъект, Отказ, ПроверяемыеРеквизиты);
	
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, РежимПроведения)
	
	Документы.КадровыйПеревод.ОбработкаПроведения(ЭтотОбъект, Отказ, РежимПроведения);
	
КонецПроцедуры

Процедура ОбработкаЗаполнения(ДанныеЗаполнения, СтандартнаяОбработка)
	
	ДатаНачала = ТекущаяДатаСеанса();
	
	Если ТипЗнч(ДанныеЗаполнения) = Тип("Структура") Тогда
		ЗаполнитьЗначенияСвойств(ЭтотОбъект, ДанныеЗаполнения);
	КонецЕсли;
	
	Если ТипЗнч(ДанныеЗаполнения) = Тип("СправочникСсылка.Сотрудники") Тогда
		ЗарплатаКадры.ЗаполнитьПоОснованиюСотрудником(ЭтотОбъект, ДанныеЗаполнения);
	КонецЕсли;
	
	Если ТипЗнч(ДанныеЗаполнения) <> Тип("Структура") Или Не ДанныеЗаполнения.Свойство("ЕжегодныеОтпуска") Тогда
		Если ТипЗнч(ДанныеЗаполнения) = Тип("Структура") Тогда
			Если ДанныеЗаполнения.Свойство("Сотрудник") Тогда
				Сотрудник						= ДанныеЗаполнения.Сотрудник;
			КонецЕсли;
			Если ДанныеЗаполнения.Свойство("ДатаНачала") Тогда
				ДатаНачала						= ДанныеЗаполнения.ДатаНачала;
			КонецЕсли;
		КонецЕсли;
	КонецЕсли;
	
	торо_ЗаполнениеДокументов.ЗаполнитьСтандартныеРеквизитыШапкиПоУмолчанию(ЭтотОбъект);
	
КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	Если ЭтоНовый() Тогда
		 торо_Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
КонецПроцедуры


#КонецОбласти

#Область СлужебныеПроцедурыИФункции
Функция ИндивидуальныйГрафикНаМесяцПеревода() Экспорт 
	ДанныеГрафика = Новый Структура("Ссылка,Дата,Номер");
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	|	ИндивидуальныйГрафик.Ссылка,
	|	ИндивидуальныйГрафик.Номер,
	|	ИндивидуальныйГрафик.Дата
	|ИЗ
	|	Документ.ИндивидуальныйГрафик.ДанныеОВремени КАК ИндивидуальныйГрафикДанныеОВремени
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ИндивидуальныйГрафик КАК ИндивидуальныйГрафик
	|		ПО ИндивидуальныйГрафикДанныеОВремени.Ссылка = ИндивидуальныйГрафик.Ссылка
	|ГДЕ
	|	ИндивидуальныйГрафикДанныеОВремени.Сотрудник = &Сотрудник
	|	И НАЧАЛОПЕРИОДА(ИндивидуальныйГрафик.ПериодРегистрации, МЕСЯЦ) = НАЧАЛОПЕРИОДА(&ДатаНачала, МЕСЯЦ)
	|	И ИндивидуальныйГрафик.Проведен
	|
	|СГРУППИРОВАТЬ ПО
	|	ИндивидуальныйГрафик.Ссылка,
	|	ИндивидуальныйГрафик.Номер,
	|	ИндивидуальныйГрафик.Дата";
	Запрос.УстановитьПараметр("ДатаНачала", НачалоМесяца(ДатаНачала));
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	Результат = Запрос.Выполнить();
	Если Не Результат.Пустой() Тогда
		Выборка = Результат.Выбрать();
		Выборка.Следующий();
		ЗаполнитьЗначенияСвойств(ДанныеГрафика, Выборка);
	КонецЕсли;
	
	Возврат ДанныеГрафика
КонецФункции

#КонецОбласти

#КонецЕсли
