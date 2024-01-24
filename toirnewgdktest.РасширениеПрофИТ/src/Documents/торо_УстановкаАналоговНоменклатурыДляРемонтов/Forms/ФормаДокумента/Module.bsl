
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #4, Башинская А.Ю., 14.08.2023
	НовыйЭлемент = Элементы.Добавить("проф_Статус", Тип("ПолеФормы"), Элементы.ГруппаПериодДействия);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_Статус";    
	НовыйЭлемент.РастягиватьПоГоризонтали = Ложь; 
		
	НовыйЭлемент = Элементы.Добавить("проф_Комментарий", Тип("ПолеФормы"), ЭтаФорма);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_Комментарий";    
	НовыйЭлемент.РастягиватьПоГоризонтали = Истина;  
	НовыйЭлемент.АвтоМаксимальнаяШирина = Ложь;
	//-- Проф-ИТ, #4, Башинская А.Ю., 14.08.2023
	
	//++ Проф-ИТ, #246, Башинская А.Ю., 10.09.2023
	НовыйЭлемент = Элементы.Добавить("проф_НадписьНомераСтрок", Тип("ДекорацияФормы"), Элементы.ГруппаОбластьДействия);
	НовыйЭлемент.Вид = ВидДекорацииФормы.Надпись;
	НовыйЭлемент.Заголовок = "Внимание! Номер строки оригинала соответствует номеру строки аналога!"; 
	НовыйЭлемент.РастягиватьПоГоризонтали = Истина;
	НовыйЭлемент.ЦветТекста = WebЦвета.Красный; 
	НовыйЭлемент.АвтоМаксимальнаяШирина = Ложь;
	//-- Проф-ИТ, #246, Башинская А.Ю., 10.09.2023      
		
КонецПроцедуры

&НаКлиенте
&ИзменениеИКонтроль("ПредставлениеОбластиДействияНажатие")
Процедура проф_ПредставлениеОбластиДействияНажатие(Элемент, СтандартнаяОбработка)

	СтандартнаяОбработка = Ложь;

	ПараметрыФормы = Новый Структура;

	ПараметрыФормы.Вставить("Организация", Объект.Организация);
	ПараметрыФормы.Вставить("ВидРемонта", Объект.ВидРемонта);
	ПараметрыФормы.Вставить("Подразделение", Объект.Подразделение);
	ПараметрыФормы.Вставить("ТехКарта", Объект.ТехКарта);
	ПараметрыФормы.Вставить("ОбъектРемонта", Объект.ОбъектРемонта);
	#Удаление
	ПараметрыФормы.Вставить("НаправлениеОбъектаРемонта", Объект.НаправлениеОбъектаРемонта);      
	#КонецУдаления  
	#Вставка
	//++ Проф-ИТ, #4, Башинская А.Ю., 28.08.2023
	ПараметрыФормы.Вставить("НаправлениеОбъектаРемонта", Объект.проф_НаправлениеДеятельности);
	#КонецВставки
	ПараметрыФормы.Вставить("СписокОбъектовРМ", Объект.СписокОбъектовРМ);
	ПараметрыФормы.Вставить("КоличествоМатериалов", Объект.Материалы.Количество());
	
	ОткрытьФорму("Документ.торо_УстановкаАналоговНоменклатурыДляРемонтов.Форма.ФормаВводаОбластиДействия", ПараметрыФормы, ЭтаФорма);

КонецПроцедуры

&НаСервере
&ИзменениеИКонтроль("ОбновитьПредставлениеОбластиДействия")
Процедура проф_ОбновитьПредставлениеОбластиДействия()

	ПредставлениеОбластиДействия = "";

	Если ЗначениеЗаполнено(Объект.Организация) Тогда

		Шаблон = НСтр("ru = 'Организация: %1;'") + " ";

		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.Организация);

	КонецЕсли;

	Если ЗначениеЗаполнено(Объект.ВидРемонта) Тогда

		Шаблон = НСтр("ru = 'Вид ремонта: %1;'") + " ";

		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.ВидРемонта);

	КонецЕсли;

	Если ЗначениеЗаполнено(Объект.Подразделение) Тогда

		Шаблон = НСтр("ru = 'Подразделение: %1;'") + " ";

		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.Подразделение);

	КонецЕсли;

	Если ЗначениеЗаполнено(Объект.ТехКарта) Тогда

		Шаблон = НСтр("ru = 'Технологическая карта: %1;'") + " ";

		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.ТехКарта);

	КонецЕсли;

	Если ЗначениеЗаполнено(Объект.ОбъектРемонта) Тогда

		Шаблон = НСтр("ru = 'Объект ремонта: %1;'") + " ";

		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.ОбъектРемонта);

	КонецЕсли;
	
	#Удаление
	Если ЗначениеЗаполнено(Объект.НаправлениеОбъектаРемонта) Тогда
		
		Шаблон = НСтр("ru = 'Направление объекта ремонта: %1;'") + " ";
		
		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.НаправлениеОбъектаРемонта);
		
	КонецЕсли;
	#КонецУдаления
	#Вставка
	//++ Проф-ИТ, #4, Башинская А.Ю., 28.08.2023
	Если ЗначениеЗаполнено(Объект.проф_НаправлениеДеятельности) Тогда
		
		Шаблон = НСтр("ru = 'Направление объекта ремонта: %1;'") + " ";
		
		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.проф_НаправлениеДеятельности);
		
	КонецЕсли;     
	//-- Проф-ИТ, #4, Башинская А.Ю., 28.08.2023
	#КонецВставки
	
	Если ЗначениеЗаполнено(Объект.СписокОбъектовРМ) Тогда 
		Шаблон = НСтр("ru = 'Список объектов регламентного мероприятия: %1;'") + " ";
		
		ПредставлениеОбластиДействия = ПредставлениеОбластиДействия + СтрШаблон(Шаблон, Объект.СписокОбъектовРМ);
		
	КонецЕсли;

	Если СтрДлина(ПредставлениеОбластиДействия) > 0 Тогда
		ПредставлениеОбластиДействия = Лев(ПредставлениеОбластиДействия, СтрДлина(ПредставлениеОбластиДействия) - 2);
	Иначе
		ПредставлениеОбластиДействия = НСтр("ru = 'По предприятию в целом'");
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
&ИзменениеИКонтроль("ОбработкаВыбора")
Процедура проф_ОбработкаВыбора(ВыбранноеЗначение, ИсточникВыбора)

	ВыполняемаяОперация = "";
	Если ТипЗнч(ВыбранноеЗначение) = Тип("Структура") И ВыбранноеЗначение.Свойство("ВыполняемаяОперация", ВыполняемаяОперация) и ВыполняемаяОперация = "ВыборОбластиДействия" Тогда

		Модифицированность = Истина;

		ЗначениеВыбора = ВыбранноеЗначение.ЗначениеВыбора;

		СписокСвойств = "Организация, ВидРемонта, Подразделение, ТехКарта, ОбъектРемонта, НаправлениеОбъектаРемонта, СписокОбъектовРМ";
		
		ЗаполнитьЗначенияСвойств(Объект, ЗначениеВыбора, СписокСвойств);
		#Вставка
		//++ Проф-ИТ, #4, Башинская А.Ю., 28.08.2023
		Объект.проф_НаправлениеДеятельности = ЗначениеВыбора.НаправлениеОбъектаРемонта;
		#КонецВставки
		ОбновитьПредставлениеОбластиДействия();

	КонецЕсли;

КонецПроцедуры

&НаСервере
Процедура проф_ОбработкаПроверкиЗаполненияНаСервереПосле(Отказ, ПроверяемыеРеквизиты)
	
	//++ Проф-ИТ, #246, Башинская А.Ю., 10.09.2023
	Если Объект.Материалы.Количество() <> Объект.Аналоги.Количество() Тогда
		ТекстСообщения = "Количество строк оригиналов не соответствует количеству строк аналогов!";
		ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,"Объект.Материалы[0].НомерСтроки",, Отказ);
		Возврат;	
	КонецЕсли;   
	
	ОтборСтроки = Новый Структура("НомерСтроки");
	Для Каждого СтрМ Из Объект.Материалы Цикл                
		ОтборСтроки.НомерСтроки = СтрМ.НомерСтроки;
		СтрА = Объект.Аналоги.НайтиСтроки(ОтборСтроки);
		Если СтрА.Количество() И СтрА[0].Номенклатура = СтрМ.Номенклатура Тогда    
			ТекстСообщения = "По строке " + СтрМ.НомерСтроки + " одинаковая номенклатура! Номенклатура оригинала не должна быть равна номенклатуре аналога.";
			ОбщегоНазначения.СообщитьПользователю(ТекстСообщения,,"Объект.Материалы[" + СтрМ.НомерСтроки + "].Номенклатура",, Отказ);
			Прервать;
		КонецЕсли;
	КонецЦикла;
	//-- Проф-ИТ, #246, Башинская А.Ю., 10.09.2023
	
КонецПроцедуры

#КонецОбласти