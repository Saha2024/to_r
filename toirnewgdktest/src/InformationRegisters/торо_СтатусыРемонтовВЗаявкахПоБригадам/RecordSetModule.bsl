
#Область ОбработчикиСобытий

Процедура ПередЗаписью(Отказ, Замещение)
	
	ЗарегистрироватьСвязанныеЗаявкиВПриложениеМобильныеБригады(Замещение, Отказ);
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ЗарегистрироватьСвязанныеЗаявкиВПриложениеМобильныеБригады(Замещение, Отказ)
	
	ЗаявкиВСоответствии = Новый Соответствие;
	
	Если Замещение Тогда
		СтруктураОтбора = Новый Структура;
		ТекстыУсловия = Новый Массив;
		Для Каждого ЭлементОтбора Из Отбор Цикл
			Если ЭлементОтбора.Использование Тогда
				СтруктураОтбора.Вставить(ЭлементОтбора.Имя, ЭлементОтбора.Значение);
				ТекстыУсловия.Добавить(СтрШаблон("%1 = &%1", ЭлементОтбора.Имя));
			КонецЕсли;
		КонецЦикла;
		
		Запрос = Новый Запрос(СтрЗаменить(
		"ВЫБРАТЬ РАЗЛИЧНЫЕ
		|	торо_СтатусыРемонтовВЗаявках.ЗаявкаНаРемонт КАК ЗаявкаНаРемонт
		|ИЗ
		|	РегистрСведений.торо_СтатусыРемонтовВЗаявкахПоБригадам КАК торо_СтатусыРемонтовВЗаявках
		|ГДЕ
		|	&Условие",
		"&Условие",
		СтрСоединить(ТекстыУсловия, " И ")
		));
		Для Каждого КлючЗначение Из СтруктураОтбора Цикл
			Запрос.УстановитьПараметр(КлючЗначение.Ключ, КлючЗначение.Значение);
		КонецЦикла;
		
		Результат = Запрос.Выполнить();
		Если Не Результат.Пустой() Тогда
			Выборка = Результат.Выбрать();
			Пока Выборка.Следующий() Цикл
				
				ЗаявкиВСоответствии.Вставить(Выборка.ЗаявкаНаРемонт, Истина);
				
			КонецЦикла;
		КонецЕсли;
		
	КонецЕсли;
	
	Для Каждого Запись Из ЭтотОбъект Цикл
		ЗаявкиВСоответствии.Вставить(Запись.ЗаявкаНаРемонт, Истина);
	КонецЦикла;
	
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	mobileBrigades.Ссылка КАК Ссылка,
	               |	mobileBrigades.ИспользоватьЗаявки КАК ИспользоватьЗаявки
	               |ИЗ
	               |	ПланОбмена.mobileBrigades КАК mobileBrigades
	               |ГДЕ
	               |	mobileBrigades.ИспользоватьЗаявки = ИСТИНА
	               |	И НЕ mobileBrigades.Ссылка = &ЭтотУзел";

	
	ЭтотУзел = ПланыОбмена.mobileBrigades.ЭтотУзел();
	Запрос.УстановитьПараметр("ЭтотУзел", ЭтотУзел);
		
	Получатели = Запрос.Выполнить().Выгрузить().ВыгрузитьКолонку("Ссылка");
		
	Для Каждого КлючЗнач Из ЗаявкиВСоответствии Цикл
		ДокОбъект = КлючЗнач.Ключ.ПолучитьОбъект();
		Если ДокОбъект = Неопределено Тогда
			Продолжить;
		КонецЕсли;  
		
		Если Получатели.Количество() > 0 Тогда
			ПланыОбмена.ЗарегистрироватьИзменения(Получатели, ДокОбъект.Ссылка);
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

#КонецОбласти