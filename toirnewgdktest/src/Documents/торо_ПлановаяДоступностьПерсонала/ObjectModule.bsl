#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область ОбработчикиСобытий

Процедура ОбработкаЗаполнения(ДанныеЗаполнения, ТекстЗаполнения, СтандартнаяОбработка)
	
	ОкончаниеПланирования = НачалоМесяца(ТекущаяДатаСеанса()) - 1;
	НачалоПланирования    = НачалоМесяца(ОкончаниеПланирования);
	
	торо_ЗаполнениеДокументов.ЗаполнитьСтандартныеРеквизитыШапкиПоУмолчанию(ЭтотОбъект);
	
КонецПроцедуры

Процедура ОбработкаПроверкиЗаполнения(Отказ, ПроверяемыеРеквизиты)
	
	// проверка на состояния сотрудника
	ПроверитьСостоянияСотрудников(Отказ);
	
	// проверить суммарное время сотрудника
	ПроверитьСуммарноеВремяРаботыСотрудников(Отказ);
	
	// проверка на пересечение периодов
	
	
КонецПроцедуры

Процедура ОбработкаПроведения(Отказ, РежимПроведения)
	
	ДвиженияПоРегистру_торо_СостояниеРаботыСотрудников(Отказ, РежимПроведения);
	
КонецПроцедуры

Процедура ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	Если ЭтоНовый() Тогда
		Автор = Пользователи.ТекущийПользователь();
	КонецЕсли;
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ДвиженияПоРегистру_торо_СостояниеРаботыСотрудников(Отказ, Режим)
	
	тзРабочееВремя = РабочееВремяСотрудников.Выгрузить();
	тзРабочееВремя.Сортировать("Сотрудник,ДатаГрафика,ВремяНачала");
	
	ТекСотрудник = Неопределено; ТекДатаГрафика = Неопределено;
	НомерПериода = 1;
	
	Для каждого СтрокаТЗ Из тзРабочееВремя Цикл
	
		Если ТекСотрудник <> СтрокаТЗ.Сотрудник ИЛИ ТекДатаГрафика <> СтрокаТЗ.ДатаГрафика Тогда
			ТекСотрудник = СтрокаТЗ.Сотрудник;
			ТекДатаГрафика = СтрокаТЗ.ДатаГрафика;
			НомерПериода = 1;
		КонецЕсли;
		//-----------------------------------
		Движение = Движения.торо_СостояниеРаботыСотрудников.Добавить();
		ЗаполнитьЗначенияСвойств(Движение, СтрокаТЗ);
		Движение.Организация    = Организация;
		Движение.Подразделение  = Подразделение;
		Движение.Период         = Дата;
		Движение.НомерПериода   = НомерПериода;
		Движение.ВремяНачала    = СоединитьДатуИВремя(СтрокаТЗ.ДатаГрафика, СтрокаТЗ.ВремяНачала);
		Движение.ВремяОкончания = СоединитьДатуИВремя(СтрокаТЗ.ДатаГрафика, СтрокаТЗ.ВремяОкончания);
		//-----------------------------------
		НомерПериода = НомерПериода + 1;
	
	КонецЦикла;
	
	ДобавитьОтмененныеПериоды(тзРабочееВремя);
	
	Движения.торо_СостояниеРаботыСотрудников.Записать(Истина);

КонецПроцедуры

Функция СоединитьДатуИВремя(Знач лДата, Знач лВремя)

	лДатаВремя = лВремя - НачалоДня(лВремя);
	Возврат НачалоДня(лДата) + лДатаВремя;

КонецФункции // СоединитьДатуИВремя()

Процедура ДобавитьОтмененныеПериоды(табРабочееВремяСотрудников)
	
	#Область ТекстЗапроса
	ТекстЗапроса = "ВЫБРАТЬ
	               |	табРабочееВремяСотрудников.Сотрудник КАК Сотрудник,
	               |	табРабочееВремяСотрудников.ДатаГрафика КАК ДатаГрафика,
	               |	табРабочееВремяСотрудников.НомерСтроки КАК НомерСтроки
	               |ПОМЕСТИТЬ ВТРабочееВремяСотрудников
	               |ИЗ
	               |	&табРабочееВремяСотрудников КАК табРабочееВремяСотрудников
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТРабочееВремяСотрудников.Сотрудник КАК Сотрудник,
	               |	ВТРабочееВремяСотрудников.ДатаГрафика КАК ДатаГрафика,
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВТРабочееВремяСотрудников.НомерСтроки) КАК НомерПериода
	               |ПОМЕСТИТЬ ВТМаксПериодСотрудникков
	               |ИЗ
	               |	ВТРабочееВремяСотрудников КАК ВТРабочееВремяСотрудников
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТРабочееВремяСотрудников.Сотрудник,
	               |	ВТРабочееВремяСотрудников.ДатаГрафика
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.Организация КАК Организация,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.Подразделение КАК Подразделение,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.Сотрудник КАК Сотрудник,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.ДатаГрафика КАК ДатаГрафика,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.НомерПериода КАК НомерПериода,
	               |	0 КАК ВремяРаботы,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.СостояниеСотрудника КАК СостояниеСотрудника,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.ВремяНачала КАК ВремяНачала,
	               |	торо_СостояниеРаботыСотрудниковСрезПоследних.ВремяОкончания КАК ВремяОкончания,
	               |	ИСТИНА КАК Отменен
	               |ИЗ
	               |	ВТМаксПериодСотрудникков КАК ВТМаксПериодСотрудникков
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_СостояниеРаботыСотрудников.СрезПоследних(
	               |				&МоментВремени,
	               |				Организация = &Организация
	               |					И Подразделение = &Подразделение) КАК торо_СостояниеРаботыСотрудниковСрезПоследних
	               |		ПО ВТМаксПериодСотрудникков.Сотрудник = торо_СостояниеРаботыСотрудниковСрезПоследних.Сотрудник
	               |			И ВТМаксПериодСотрудникков.ДатаГрафика = торо_СостояниеРаботыСотрудниковСрезПоследних.ДатаГрафика
	               |			И ВТМаксПериодСотрудникков.НомерПериода < торо_СостояниеРаботыСотрудниковСрезПоследних.НомерПериода";
	#КонецОбласти
	
	Запрос = Новый Запрос(ТекстЗапроса);
	Запрос.УстановитьПараметр("МоментВремени", Новый МоментВремени(Дата, Ссылка));
	Запрос.УстановитьПараметр("Организация",   Организация);
	Запрос.УстановитьПараметр("Подразделение", Подразделение);
	Запрос.УстановитьПараметр("табРабочееВремяСотрудников", табРабочееВремяСотрудников);
	
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
	
		Движение = Движения.торо_СостояниеРаботыСотрудников.Добавить();
		Движение.Период = Дата;
		ЗаполнитьЗначенияСвойств(Движение,Выборка);
	
	КонецЦикла; 
	
КонецПроцедуры

Процедура ПроверитьСостоянияСотрудников(Отказ)

	#Область ТекстЗапроса
	ТекстЗапроса = "ВЫБРАТЬ
	               |	табРабочееВремяСотрудников.Сотрудник КАК Сотрудник,
	               |	табРабочееВремяСотрудников.ДатаГрафика КАК ДатаГрафика,
	               |	табРабочееВремяСотрудников.СостояниеСотрудника КАК СостояниеСотрудника,
	               |	табРабочееВремяСотрудников.НомерСтроки КАК НомерСтроки
	               |ПОМЕСТИТЬ ВТРабВремяСотрудников
	               |ИЗ
	               |	&табРабочееВремяСотрудников КАК табРабочееВремяСотрудников
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТРабВремяСотрудников.Сотрудник КАК Сотрудник,
	               |	ВТРабВремяСотрудников.ДатаГрафика КАК ДатаГрафика,
	               |	МАКСИМУМ(ВТРабВремяСотрудников.НомерСтроки) КАК НомерСтроки
	               |ИЗ
	               |	ВТРабВремяСотрудников КАК ВТРабВремяСотрудников
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТРабВремяСотрудников.Сотрудник,
	               |	ВТРабВремяСотрудников.ДатаГрафика
	               |
	               |ИМЕЮЩИЕ
	               |	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВТРабВремяСотрудников.СостояниеСотрудника) > 1";
	#КонецОбласти 
	
	Запрос = Новый Запрос(ТекстЗапроса);
	Запрос.УстановитьПараметр("табРабочееВремяСотрудников", РабочееВремяСотрудников);
	
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
	
		ТекстСообщения = СтрШаблон("Для сотрудника %1 на %2 указаны различные состояния (строка %3)", Выборка.Сотрудник, Формат(Выборка.ДатаГрафика,"ДФ=dd.MM.yy") , Выборка.НомерСтроки);
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(ТекстСообщения,Ссылка,,,Отказ);
	
	КонецЦикла;

КонецПроцедуры

Процедура ПроверитьСуммарноеВремяРаботыСотрудников(Отказ)
	
	#Область ТекстЗапроса
	ТекстЗапроса = "ВЫБРАТЬ
	               |	табРабочееВремяСотрудников.Сотрудник КАК Сотрудник,
	               |	табРабочееВремяСотрудников.ДатаГрафика КАК ДатаГрафика,
	               |	табРабочееВремяСотрудников.ВремяРаботы КАК ВремяРаботы,
	               |	табРабочееВремяСотрудников.НомерСтроки КАК НомерСтроки
	               |ПОМЕСТИТЬ ВТРабВремяСотрудников
	               |ИЗ
	               |	&табРабочееВремяСотрудников КАК табРабочееВремяСотрудников
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТРабВремяСотрудников.Сотрудник КАК Сотрудник,
	               |	ВТРабВремяСотрудников.ДатаГрафика КАК ДатаГрафика,
	               |	МАКСИМУМ(ВТРабВремяСотрудников.НомерСтроки) КАК НомерСтроки
	               |ИЗ
	               |	ВТРабВремяСотрудников КАК ВТРабВремяСотрудников
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТРабВремяСотрудников.Сотрудник,
	               |	ВТРабВремяСотрудников.ДатаГрафика
	               |
	               |ИМЕЮЩИЕ
	               |	СУММА(ВТРабВремяСотрудников.ВремяРаботы) > &ЛимитВремениРаботы";
	#КонецОбласти 
	
	Запрос = Новый Запрос(ТекстЗапроса);
	Запрос.УстановитьПараметр("табРабочееВремяСотрудников", РабочееВремяСотрудников);
	Запрос.УстановитьПараметр("ЛимитВремениРаботы", 24*3600);   // 24 часа
	
	Выборка = Запрос.Выполнить().Выбрать();
	Пока Выборка.Следующий() Цикл
	
		ТекстСообщения = СтрШаблон("Для сотрудника %1 на %2 суммарное время работы превышает 24 часа (строка %3)", Выборка.Сотрудник, Формат(Выборка.ДатаГрафика,"ДФ=dd.MM.yy") , Выборка.НомерСтроки);
		ОбщегоНазначенияКлиентСервер.СообщитьПользователю(ТекстСообщения,Ссылка,,,Отказ);
	
	КонецЦикла;

КонецПроцедуры

#КонецОбласти 

#КонецЕсли