#Область ПрограммныйИнтерфейс

&ИзменениеИКонтроль("УстановитьСтатусРемонтовПриПроведении")
Процедура проф_УстановитьСтатусРемонтовПриПроведении(Источник)

	ДокументСсылка = Источник.Ссылка;
	ВремяУстановки = Источник.Дата;

	Статус = Перечисления.торо_СтатусыРемонтов.Зарегистрирован;
	Если ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ВнешнееОснованиеДляРабот") Тогда
		ТабИДРемонтов = Источник.ОбследованноеОборудование.Выгрузить(, "ID");
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ВыявленныеДефекты") Тогда
		ТабИДРемонтов = Источник.СписокДефектов.Выгрузить(, "ID");
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ПланГрафикРемонта") 
		ИЛИ ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ГрафикРегламентныхМероприятийТОиР") Тогда
		Если ДокументСсылка.ВидОперации = Перечисления.торо_ВидыОперацийПланаГрафикаППР.Корректировка Тогда
			ТабИДРемонтов = Источник.ПланРемонтов.Выгрузить(Новый Структура("Отменен", Истина), "ID");
			Статус = Перечисления.торо_СтатусыРемонтов.Отменен;
		Иначе
			ТабИДРемонтов = Источник.ПланРемонтов.Выгрузить(, "ID");
		КонецЕсли;
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") Тогда
		ТабИДРемонтов = Источник.РемонтыОборудования.Выгрузить(, "ID");
		Статус = Перечисления.торо_СтатусыРемонтов.ЗапланированоУстранение;
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ЗакрытиеЗаявокИРемонтов") Тогда
		ТабИДРемонтов = Источник.ЗакрываемыеРемонты.Выгрузить(, "ID, Перенесен");
		Статус = Перечисления.торо_СтатусыРемонтов.Отменен;		
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_НарядНаРегламентноеМероприятие") Тогда
		ТабИДРемонтов = Источник.РегламентныеМероприятия.Выгрузить(, "ID");
		Статус = Перечисления.торо_СтатусыРемонтов.ВРаботе;
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_АктОВыполненииРегламентногоМероприятия") Тогда
		ТабИДРемонтов = Источник.Мероприятия.Выгрузить(, "ID, ЗавершитьРемонт");
		Статус = Перечисления.торо_СтатусыРемонтов.ВРаботе;
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_НарядНаВыполнениеРемонтныхРабот") Тогда 
		ТабИДРемонтов = Источник.РемонтыОборудования.Выгрузить(, "ID");
		Статус = Перечисления.торо_СтатусыРемонтов.ВРаботе;
	ИначеЕсли ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_АктОВыполненииЭтапаРабот") Тогда
		ТабИДРемонтов = Источник.РемонтыОборудования.Выгрузить(, "ID, ЗавершитьРемонт");
		Статус = Перечисления.торо_СтатусыРемонтов.ВРаботе;
	КонецЕсли;

	УстановитьУправляемыеБлокировкиНаСтатусыРемонтов(ДокументСсылка);

	СоответствиеТекущихСтатусов = ДанныеСтатусовРемонтов(ТабИДРемонтов);

	ПараметрыДляПроверки = ПолучитьСтруктуруПараметровДляПроверкиСменыСтатусаРемонта();
	ПараметрыДляПроверки.Документ = ДокументСсылка;
	ПараметрыДляПроверки.ДокументИсточник = Источник.Ссылка;
	ПараметрыДляПроверки.ПроверятьСтатусДокумента = Ложь;
	ПараметрыДляПроверки.ПолучатьДанныеТекущегоСтатуса = Ложь;

	НЗСтатусыРемонта = РегистрыСведений.торо_СтатусыРемонтов.СоздатьНаборЗаписей();
	НЗСтатусыРемонта.Отбор.Регистратор.Установить(ДокументСсылка);
	НЗСтатусыРемонта.Прочитать();

	УдалитьЗаписиСтатусовВНабореЗаписейПоДокументуИсточнику(НЗСтатусыРемонта, Источник.Ссылка);

	#Вставка
	//++ Проф-ИТ, #228, Соловьев А.А., 04.10.2023
	Если ТипЗнч(ДокументСсылка) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") Тогда 
		ЭтоЗаявкаНаРемонтКВыполнению = Источник.проф_КВыполнению;
	Иначе
		ЭтоЗаявкаНаРемонтКВыполнению = Ложь;
	КонецЕсли;	
	//-- Проф-ИТ, #228, Соловьев А.А., 04.10.2023
	#КонецВставки
	Для Каждого Строка Из ТабИДРемонтов Цикл

		ДанныеТекущегоСтатуса = СоответствиеТекущихСтатусов[Строка.ID];

		УстанавливаемыйСтатус = Статус;
		Если ТабИДРемонтов.Колонки.Найти("Перенесен") <> Неопределено И Строка.Перенесен Тогда
			УстанавливаемыйСтатус = Перечисления.торо_СтатусыРемонтов.Зарегистрирован;
		КонецЕсли;
		Если ТабИДРемонтов.Колонки.Найти("ЗавершитьРемонт") <> Неопределено И Строка.ЗавершитьРемонт Тогда
			УстанавливаемыйСтатус = Перечисления.торо_СтатусыРемонтов.Выполнен;
		КонецЕсли;

		ПараметрыДляПроверки.ИДРемонта = Строка.ID;
		ПараметрыДляПроверки.УстанавливаемыйСтатус = УстанавливаемыйСтатус;
		ПараметрыДляПроверки.ДанныеТекущегоСтатуса = ДанныеТекущегоСтатуса;

		МожноСменитьСтатус = ПроверитьВозможностьУстановкиСтатусаДляРемонта(ПараметрыДляПроверки, ТабИДРемонтов);

		#Вставка
		//++ Проф-ИТ, #228, Соловьев А.А., 04.10.2023
		Если ЭтоЗаявкаНаРемонтКВыполнению Тогда
			УстанавливаемыйСтатус = Перечисления.торо_СтатусыРемонтов.проф_КВыполнению;
			МожноСменитьСтатус = Истина;
		КонецЕсли;	
		//-- Проф-ИТ, #228, Соловьев А.А., 04.10.2023
		#КонецВставки
		Если МожноСменитьСтатус Тогда
			Запись = НЗСтатусыРемонта.Добавить();
			Запись.Регистратор = ДокументСсылка;
			Запись.IDРемонта = Строка.ID;
			Запись.Период = ВремяУстановки;
			Запись.Статус = УстанавливаемыйСтатус;
		КонецЕсли;

	КонецЦикла;

	НЗСтатусыРемонта.Записать(Истина);

	ЗаявкиНаРемонт = ПолучитьЗаявки(ДокументСсылка);
	Для каждого Заявка Из ЗаявкиНаРемонт Цикл
		УстановитьСтатусыДляИсполнителейПриПроведенииДокумента(Источник, Заявка, Статус, СоответствиеТекущихСтатусов);
	КонецЦикла;

КонецПроцедуры
	
#КонецОбласти