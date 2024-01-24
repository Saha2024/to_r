
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПосле(Отказ, СтандартнаяОбработка)
	
	//++ Проф-ИТ, #169, Башинская А.Ю., 09.08.2023
	НовыйЭлемент = Элементы.Добавить("проф_ТипЗаказа", Тип("ПолеФормы"), Элементы.ГруппаПравая);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_ТипЗаказа";
	//++ Проф-ИТ, #358, Соловьев А.А., 20.11.2023
	НовыйЭлемент.УстановитьДействие("ПриИзменении", "проф_ТипЗаказаПриИзменении");
	//-- Проф-ИТ, #358, Соловьев А.А., 20.11.2023
	
	НовыйЭлемент = Элементы.Добавить("проф_КрайнееВремяУстранения", Тип("ПолеФормы"), Элементы.ГруппаДатаОбнаружения);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_КрайнееВремяУстранения";     
		
	НовыйЭлемент = Элементы.Добавить("проф_РемБаза", Тип("ПолеФормы"), Элементы.ГруппаПравая);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_РемБаза";         
	
	Если Параметры.Ключ.Пустая() Тогда   
		Объект.проф_ТипЗаказа = Перечисления.проф_ТипЗаказаНаРемонт.Плановый;
	КонецЕсли;	
	//-- Проф-ИТ, #169, Башинская А.Ю., 09.08.2023  
	
	//++ Проф-ИТ, #133, Антонов Б.Л., 20.09.2023
	НовыйЭлементГруппа           = Элементы.Вставить("проф_ГруппаПодписанты", Тип("ГруппаФормы"), Элементы.Группа5, Элементы.ГруппаДополнительныеРеквизиты);
	НовыйЭлементГруппа.Вид       = ВидГруппыФормы.Страница;
	НовыйЭлементГруппа.Заголовок = "Подписанты";
	
	НовыйЭлементТаблица = Элементы.Добавить("проф_Подписанты", Тип("ТаблицаФормы"), НовыйЭлементГруппа);
	НовыйЭлементТаблица.ПутьКДанным = "Объект.проф_Подписанты"; 

	ЭлементТаблицыФормы     = Элементы.Добавить("проф_ПодписантыРоль", Тип("ПолеФормы"), НовыйЭлементТаблица);
	ЭлементТаблицыФормы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицыФормы.ПутьКДанным = "Объект.проф_Подписанты.Роль";
	
	ЭлементТаблицыФормы     = Элементы.Добавить("проф_ПодписантыСотрудник", Тип("ПолеФормы"), НовыйЭлементТаблица);
	ЭлементТаблицыФормы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицыФормы.ПутьКДанным = "Объект.проф_Подписанты.Сотрудник";
	ЭлементТаблицыФормы.УстановитьДействие("ПриИзменении", "проф_ПодписантыСотрудникПриИзменении");

	ЭлементТаблицыФормы     = Элементы.Добавить("проф_ПодписантыДолжность", Тип("ПолеФормы"), НовыйЭлементТаблица);
	ЭлементТаблицыФормы.Вид = ВидПоляФормы.ПолеВвода;
	ЭлементТаблицыФормы.ПутьКДанным = "Объект.проф_Подписанты.Должность";
	//-- Проф-ИТ, #133, Антонов Б.Л., 20.09.2023 
	
	//++ Проф-ИТ, #326, Соловьев А.А., 31.10.2023
	НовыйЭлемент = Элементы.Добавить("проф_ЭксплуатацияВозможна", Тип("ПолеФормы"), Элементы.ГруппаЛевая);
	НовыйЭлемент.Вид = ВидПоляФормы.ПолеПереключателя;
	НовыйЭлемент.ВидПереключателя = ВидПереключателя.Тумблер;
	НовыйЭлемент.ПутьКДанным = "Объект.проф_ЭксплуатацияВозможна";
	НовыйЭлемент.СписокВыбора.Добавить("Да", "Да");
	НовыйЭлемент.СписокВыбора.Добавить("Нет", "Нет");
	Элементы.Переместить(НовыйЭлемент, Элементы.ГруппаЛевая, Элементы.Декорация1);
	//-- Проф-ИТ, #326, Соловьев А.А., 31.10.2023     
	
	//++ Проф-ИТ, #430, Сергеев Д.Н., 18.01.2024
	НовыйЭлементГруппа = Элементы.Вставить("проф_ГруппаОсновное", Тип("ГруппаФормы"), Элементы.Группа5, Элементы.ДефектыОборудования);
	НовыйЭлементГруппа.Вид       = ВидГруппыФормы.Страница;
	НовыйЭлементГруппа.Заголовок = "Основное";  
	
	НовыйГруппаERP = Элементы.Добавить("проф_ГруппаERP", Тип("ГруппаФормы"), Элементы.проф_ГруппаОсновное);		
	НовыйГруппаERP.Вид       = ВидГруппыФормы.ОбычнаяГруппа;
	НовыйГруппаERP.Заголовок = "Из ERP";
	НовыйГруппаERP.Группировка = ГруппировкаПодчиненныхЭлементовФормы.Вертикальная; 
	НовыйГруппаERP.ТолькоПросмотр = Истина;
	
	НовыйЭлемент = Элементы.Добавить("проф_ВидДефектаERP", Тип("ПолеФормы"), НовыйГруппаERP);
	НовыйЭлемент.Вид         = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.Заголовок   = "Вид дефекта";
	НовыйЭлемент.ПутьКДанным = "Объект.проф_ВидДефекта";
	
	НовыйЭлемент = Элементы.Добавить("проф_КраткоеОписаниеДефекта", Тип("ПолеФормы"), НовыйГруппаERP);
	НовыйЭлемент.Вид         = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.Заголовок   = "Краткое описание дефекта";
	НовыйЭлемент.ПутьКДанным = "Объект.проф_КраткоеОписаниеДефекта";
	НовыйЭлемент.АвтоМаксимальнаяШирина = Ложь;
	
	НовыйЭлемент = Элементы.Добавить("проф_ПолноеОписаниеДефекта", Тип("ПолеФормы"), НовыйГруппаERP);
	НовыйЭлемент.Вид         = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.Заголовок   = "Полное описание дефекта";
	НовыйЭлемент.ПутьКДанным = "Объект.проф_ПолноеОписаниеДефекта";
	НовыйЭлемент.АвтоМаксимальнаяШирина = Ложь;
	
	НовыйЭлемент = Элементы.Добавить("проф_РешениеПоУстранению", Тип("ПолеФормы"), НовыйГруппаERP);
	НовыйЭлемент.Вид         = ВидПоляФормы.ПолеВвода;
	НовыйЭлемент.Заголовок   = "Решение по устранению";
	НовыйЭлемент.ПутьКДанным = "Объект.проф_РешениеПоУстранению"; 
	НовыйЭлемент.АвтоМаксимальнаяШирина = Ложь; 
	
	Элементы.Переместить(Элементы.КонтактноеЛицоИнициатора, НовыйЭлементГруппа);
	Элементы.Переместить(Элементы.Инициатор, НовыйЭлементГруппа);
	Элементы.Переместить(Элементы.ПодразделениеИсполнитель, НовыйЭлементГруппа);
	Элементы.Переместить(Элементы.проф_ТипЗаказа, НовыйЭлементГруппа);
	Элементы.Переместить(Элементы.проф_РемБаза, НовыйЭлементГруппа);
	//-- Проф-ИТ, #430, Сергеев Д.Н., 18.01.2024	
	
	//++ Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	проф_ОбщегоНазначенияВызовСервера.УстановитьПараметрВыбораПодразделения(Элементы);
	//-- Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

//++ Проф-ИТ, #133, Антонов Б.Л., 20.09.2023
&НаКлиенте
Процедура проф_ПодписантыСотрудникПриИзменении()
	
	проф_Подписанты = Элементы.проф_Подписанты.ТекущиеДанные; 
	проф_Подписанты.Должность = проф_ПодписантыСотрудникПриИзмененииНаСервере(проф_Подписанты.Сотрудник);
	
КонецПроцедуры 
//-- Проф-ИТ, #133, Антонов Б.Л., 20.09.2023 

//++ Проф-ИТ, #358, Соловьев А.А., 20.11.2023
&НаКлиенте
Процедура проф_ТипЗаказаПриИзменении(Элемент)
	
	ЗначениеДефектВыявленВПроцессе = ЗначениеДефектВыявленВПроцессе(Объект.проф_ТипЗаказа);
	Если ЗначениеДефектВыявленВПроцессе <> Неопределено Тогда 
		Объект.ДефектВыявленВПроцессе = ЗначениеДефектВыявленВПроцессе;
	КонецЕсли;
	
КонецПроцедуры
//-- Проф-ИТ, #358, Соловьев А.А., 20.11.2023

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыСписокДефектов

//++ Проф-ИТ, #64, Карпов Д.Ю., 31.08.2023
&НаКлиенте
&Вместо("СписокДефектовПриАктивизацииСтроки")
Процедура проф_СписокДефектовПриАктивизацииСтроки(Элемент)
	
	ТекущиеДанные = Элементы.СписокДефектов.ТекущиеДанные;
	
	Если ТекущиеДанные = Неопределено Тогда 
		Элементы.СписокДефектовПоказатьМестоположение.Доступность = Ложь;
	Иначе
		Если ЗначениеЗаполнено(ТекущиеДанные.ОтказавшийЭлемент) Тогда
			Элементы.СписокДефектовПоказатьМестоположение.Доступность = торо_ОбщегоНазначенияВызовСервера.ЗначениеРеквизитаОбъекта(ТекущиеДанные.ОтказавшийЭлемент, "ЭтоПротяженныйОбъект");
		Иначе 
			Элементы.СписокДефектовПоказатьМестоположение.Доступность = Ложь;
		КонецЕсли;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура проф_СписокДефектовПередУдалениемПеред(Элемент, Отказ) //++ Проф-ИТ, #430, Сергеев Д.Н., 18.01.2024
	
	текДанные = Элементы.СписокДефектов.ТекущиеДанные;
	проф_ПроверитьIDРемонтВРегистреВсеДокументыПоРемонтам(текДанные.ID, Отказ);

КонецПроцедуры //-- Проф-ИТ, #430, Сергеев Д.Н., 18.01.2024

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #133, Антонов Б.Л., 20.09.2023
&НаСервереБезКонтекста
Функция проф_ПодписантыСотрудникПриИзмененииНаСервере(Сотрудник)
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	ТекущиеКадровыеДанныеСотрудников.Сотрудник КАК Сотрудник,
		|	ТекущиеКадровыеДанныеСотрудников.ТекущаяДолжность КАК ТекущаяДолжность
		|ИЗ
		|	Справочник.Сотрудники КАК Сотрудники
		|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ТекущиеКадровыеДанныеСотрудников КАК ТекущиеКадровыеДанныеСотрудников
		|		ПО Сотрудники.Ссылка = ТекущиеКадровыеДанныеСотрудников.Сотрудник
		|ГДЕ
		|	Сотрудники.Ссылка = &Сотрудник";
	
	Запрос.УстановитьПараметр("Сотрудник", Сотрудник);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Если ВыборкаДетальныеЗаписи.Следующий() Тогда
		Возврат ВыборкаДетальныеЗаписи.ТекущаяДолжность;
	КонецЕсли;
	
	Возврат Справочники.Должности.ПустаяСсылка();
	
КонецФункции
//-- Проф-ИТ, #133, Антонов Б.Л., 20.09.2023 

//++ Проф-ИТ, #358, Соловьев А.А., 20.11.2023
&НаСервереБезКонтекста
Функция ЗначениеДефектВыявленВПроцессе(проф_ТипЗаказа)
	
	Если проф_ТипЗаказа = Перечисления.проф_ТипЗаказаНаРемонт.Плановый Тогда
		Возврат Справочники.проф_НастройкиСистемы.ПолучитьНастройкуСистемы("ЗаполнениеДокументаВыявленныйДефект", 
			"УсловиеВыявленияДефектаПлановыйОсмотр");
	ИначеЕсли проф_ТипЗаказа = Перечисления.проф_ТипЗаказаНаРемонт.Внеплановый Тогда
		Возврат Справочники.проф_НастройкиСистемы.ПолучитьНастройкуСистемы("ЗаполнениеДокументаВыявленныйДефект", 
			"УсловиеВыявленияДефектаЭксплуатация");
	Иначе
		Возврат Неопределено;
	КонецЕсли;
	
КонецФункции

//-- Проф-ИТ, #358, Соловьев А.А., 20.11.2023

&НаСервере
Процедура проф_ПроверитьIDРемонтВРегистреВсеДокументыПоРемонтам(IDРемонта, Отказ) //++ Проф-ИТ, #430, Сергеев Д.Н., 18.01.2024
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	торо_ВсеДокументыПоРемонтам.IDРемонта КАК IDРемонта,
		|	ПРЕДСТАВЛЕНИЕССЫЛКИ(торо_ВсеДокументыПоРемонтам.ДокументПоРемонту) КАК ДокументПоРемонтуПредставление
		|ИЗ
		|	РегистрСведений.торо_ВсеДокументыПоРемонтам КАК торо_ВсеДокументыПоРемонтам
		|ГДЕ
		|	торо_ВсеДокументыПоРемонтам.IDРемонта = &IDРемонта
		|	И (ТИПЗНАЧЕНИЯ(торо_ВсеДокументыПоРемонтам.ДокументПоРемонту) = ТИП(Документ.торо_ЗаявкаНаРемонт)
		|			ИЛИ ТИПЗНАЧЕНИЯ(торо_ВсеДокументыПоРемонтам.ДокументПоРемонту) = ТИП(Документ.торо_ЗакрытиеЗаявокИРемонтов)
		|			ИЛИ ТИПЗНАЧЕНИЯ(торо_ВсеДокументыПоРемонтам.ДокументПоРемонту) = ТИП(Документ.торо_НарядНаВыполнениеРемонтныхРабот)
		|			ИЛИ ТИПЗНАЧЕНИЯ(торо_ВсеДокументыПоРемонтам.ДокументПоРемонту) = ТИП(Документ.торо_АктОВыполненииЭтапаРабот))";
	
	Запрос.УстановитьПараметр("IDРемонта", IDРемонта);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	Если Не РезультатЗапроса.Пустой() Тогда
		
		Выборка = РезультатЗапроса.Выбрать();
		Пока Выборка.Следующий() Цикл 
			ТекстСообщения = НСтр("ru = 'Нельзя удалить строку, т.к. в документе %1 есть строки, которые относятся к удаляемой строке ремонта.'");
			ТекстСообщения = СтрШаблон(ТекстСообщения, Выборка.ДокументПоРемонтуПредставление);
			
			ОбщегоНазначенияКлиентСервер.СообщитьПользователю(ТекстСообщения,,,, Отказ);
		КонецЦикла;		
		
	КонецЕсли;		
	
КонецПроцедуры //-- Проф-ИТ, #430, Сергеев Д.Н., 18.01.2024

#КонецОбласти
