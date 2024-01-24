#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда
	
////////////////////////////////////////////////////////////////////////////////
// СЛУЖЕБНЫЕ ПЕРЕМЕННЫЕ

перем СтруктураДанных Экспорт;  // Структура, хранящая данные для работы с уведомлениями.
Перем БезусловнаяЗапись Экспорт; // Отключает проверки при записи документа

#Область ОбработчикиСобытий

Процедура ОбработкаЗаполнения(ДанныеЗаполнения, ТекстЗаполнения, СтандартнаяОбработка)
	
	Если ТипЗнч(ДанныеЗаполнения) = Тип("Структура") Тогда
		ЗаполнитьЗначенияСвойств(ЭтотОбъект, ДанныеЗаполнения);
	КонецЕсли;

	торо_ЗаполнениеДокументов.ЗаполнитьСтандартныеРеквизитыШапкиПоУмолчанию(ЭтотОбъект);
	
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, Режим)
		
	Движения.торо_НастройкаРегистрацииНаработкиЗависимыхОР.Записывать = Истина;
	Движения.торо_НастройкаРегистрацииНаработкиЗависимыхОР.Очистить();
	
	ТаблицаЗависимыхОбъектовРемонта = ПолучитьТаблицуЗависимыхОбъектовРемонта(ОбъектыРемонта,ЗависимыеОбъектыРемонта);
	
	Для Каждого ТекСтрокаЗависимыеОбъектыРемонта Из ТаблицаЗависимыхОбъектовРемонта Цикл
		Движение = Движения.торо_НастройкаРегистрацииНаработкиЗависимыхОР.Добавить();
		
		Движение.Период = ТекСтрокаЗависимыеОбъектыРемонта.ДатаВвода;
		Движение.Дата = ТекСтрокаЗависимыеОбъектыРемонта.ДатаВвода;
		Движение.ОбъектРемонта = ТекСтрокаЗависимыеОбъектыРемонта.ОбъектРемонта;
		Движение.ОрганизацияЗависимогоОР = ТекСтрокаЗависимыеОбъектыРемонта.ОрганизацияЗависимогоОР;
		Движение.ПодразделениеЗависимогоОР = ТекСтрокаЗависимыеОбъектыРемонта.ПодразделениеЗависимогоОР;
		Движение.НаправлениеЗависимогоОР = ТекСтрокаЗависимыеОбъектыРемонта.НаправлениеЗависимогоОР;
		Движение.ТипОРЗависимогоОбъектаРемонта = ТекСтрокаЗависимыеОбъектыРемонта.ТипОРЗависимогоОбъектаРемонта;
		Движение.ЗависимыйОбъектРемонта = ТекСтрокаЗависимыеОбъектыРемонта.ЗависимыйОбъектРемонта;
		Движение.ПараметрНаработки = ТекСтрокаЗависимыеОбъектыРемонта.ПараметрНаработки;
		Движение.КоэффициентПересчета = ТекСтрокаЗависимыеОбъектыРемонта.КоэффициентПересчета;
	КонецЦикла;

КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
КонецПроцедуры
#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Функция ПолучитьТаблицуЗависимыхОбъектовРемонта(пОбъектыРемонта,пЗависимыеОбъектыРемонта)
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	ОбъектыРемонта.ID_ОбъектаРемонта КАК ID,
	|	ОбъектыРемонта.ОбъектРемонта КАК ОбъектРемонта
	|ПОМЕСТИТЬ ОбъектыРемонта
	|ИЗ
	|	&ОбъектыРемонта КАК ОбъектыРемонта
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ЗависимыеОбъектыРемонта.ID_ВышестоящегоОбъекта КАК ID_ВышестоящегоОбъекта,
	|	ЗависимыеОбъектыРемонта.ОбъектРемонта КАК ОбъектРемонта,
	|	ЗависимыеОбъектыРемонта.ПараметрНаработки КАК ПараметрНаработки,
	|	ЗависимыеОбъектыРемонта.Организация КАК ОрганизацияЗависимогоОР,
	|	ЗависимыеОбъектыРемонта.Подразделение КАК ПодразделениеЗависимогоОР,
	|	ЗависимыеОбъектыРемонта.Направление КАК НаправлениеЗависимогоОР,
	|	ЗависимыеОбъектыРемонта.ТипОбъектаРемонта КАК ТипОРЗависимогоОбъектаРемонта,
	|	ЗависимыеОбъектыРемонта.КоэффициентПересчета КАК КоэффициентПересчета,
	|	ЗависимыеОбъектыРемонта.ДатаВвода КАК ДатаВвода
	|ПОМЕСТИТЬ ЗависимыеОбъектыРемонта
	|ИЗ
	|	&ЗависимыеОбъектыРемонта КАК ЗависимыеОбъектыРемонта
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	ОбъектыРемонта.ОбъектРемонта КАК ОбъектРемонта,
	|	ЗависимыеОбъектыРемонта.ОрганизацияЗависимогоОР КАК ОрганизацияЗависимогоОР,
	|	ЗависимыеОбъектыРемонта.ПодразделениеЗависимогоОР КАК ПодразделениеЗависимогоОР,
	|	ЗависимыеОбъектыРемонта.НаправлениеЗависимогоОР КАК НаправлениеЗависимогоОР,
	|	ЗависимыеОбъектыРемонта.ТипОРЗависимогоОбъектаРемонта КАК ТипОРЗависимогоОбъектаРемонта,
	|	ЗависимыеОбъектыРемонта.ОбъектРемонта КАК ЗависимыйОбъектРемонта,
	|	ЗависимыеОбъектыРемонта.ПараметрНаработки КАК ПараметрНаработки,
	|	ЗависимыеОбъектыРемонта.КоэффициентПересчета КАК КоэффициентПересчета,
	|	ЗависимыеОбъектыРемонта.ДатаВвода КАК ДатаВвода
	|ИЗ
	|	ЗависимыеОбъектыРемонта КАК ЗависимыеОбъектыРемонта
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ОбъектыРемонта КАК ОбъектыРемонта
	|		ПО ЗависимыеОбъектыРемонта.ID_ВышестоящегоОбъекта = ОбъектыРемонта.ID";
					
	Запрос.УстановитьПараметр("ОбъектыРемонта",пОбъектыРемонта);
	Запрос.УстановитьПараметр("ЗависимыеОбъектыРемонта",пЗависимыеОбъектыРемонта);
	Возврат Запрос.Выполнить().Выгрузить();
	
КонецФункции

#КонецОбласти

#КонецЕсли