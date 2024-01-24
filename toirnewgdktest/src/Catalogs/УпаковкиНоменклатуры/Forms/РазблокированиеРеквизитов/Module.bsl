#Область ОбработчикиКомандФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Если Параметры.Свойство("АвтоТест") Тогда // Возврат при получении формы для анализа.
		Возврат;
	КонецЕсли;
	
	РазрешитьРедактированиеКоэффициент = Истина;
	РазрешитьРедактированиеЕдиницаИзмерения = Истина;
	РазрешитьРедактированиеРодитель = Истина;
	РазрешитьРедактированиеКоличествоУпаковок = Истина;
	РазрешитьРедактированиеСостоитИзДругихУпаковок = Истина;
	
	СсылкаНаОбъект = Параметры.Ссылка;
	Элементы.ГруппаОбъектИспользуетсяСтраницы.ТекущаяСтраница = Элементы.ГруппаОбъектИспользуетсяНужноПроверить;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийФормы

&НаКлиенте
Процедура РазрешитьРедактирование(Команда)
	
	Результат = Новый Массив();
	
	Если РазрешитьРедактированиеКоэффициент Тогда
		Результат.Добавить("Коэффициент");
	КонецЕсли;
	
	Если РазрешитьРедактированиеЕдиницаИзмерения Тогда
		Результат.Добавить("ЕдиницаИзмерения");
	КонецЕсли;
	
	Если РазрешитьРедактированиеРодитель Тогда
		Результат.Добавить("Родитель");
	КонецЕсли;
	
	Если РазрешитьРедактированиеКоличествоУпаковок Тогда
		Результат.Добавить("КоличествоУпаковок");
	КонецЕсли;
	
	Если РазрешитьРедактированиеСостоитИзДругихУпаковок Тогда
		Результат.Добавить("СостоитИзДругихУпаковок");
	КонецЕсли;
	
	Закрыть(Результат);

КонецПроцедуры

&НаКлиенте
Процедура ПроверитьИспользованиеОбъекта(Команда)
	
	ЕстьСсылкиНаОбъект = ОбщегоНазначенияВызовСервера.ЕстьСсылкиНаОбъект(СсылкаНаОбъект);
	Если ЕстьСсылкиНаОбъект Тогда
		Элементы.ГруппаОбъектИспользуетсяСтраницы.ТекущаяСтраница = Элементы.ГруппаОбъектИспользуетсяОбъектИспользуется;
		Элементы.ГруппаОбъектИспользуетсяНужноПроверить.Видимость = Ложь;
	Иначе
		Элементы.ГруппаОбъектИспользуетсяСтраницы.ТекущаяСтраница = Элементы.ГруппаОбъектИспользуетсяОбъектНеИспользуется;
		Элементы.ГруппаОбъектИспользуетсяНужноПроверить.Видимость = Ложь;
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура РазрешитьРедактированиеСостоитИзДругихУпаковокПриИзменении(Элемент)
	
	Элементы.РазрешитьРедактированиеРодитель.Доступность = РазрешитьРедактированиеСостоитИзДругихУпаковок;
	Элементы.РазрешитьРедактированиеКоличествоУпаковок.Доступность = РазрешитьРедактированиеСостоитИзДругихУпаковок;
	Элементы.РазрешитьРедактированиеКоэффициент.Доступность = РазрешитьРедактированиеСостоитИзДругихУпаковок;
	
	Если Не РазрешитьРедактированиеСостоитИзДругихУпаковок = Истина Тогда
		РазрешитьРедактированиеРодитель = Ложь;
		РазрешитьРедактированиеКоличествоУпаковок = Ложь;
		РазрешитьРедактированиеКоэффициент = Ложь;
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти