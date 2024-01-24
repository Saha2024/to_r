#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	ТекСтруктураИерархии = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить(
		"НастройкиТОиР",
		"ОсновнаяСтруктураИерархии",
		Истина);
		
	ПоследниеШтрихКоды = Новый Структура;
	ПоследниеШтрихКоды.Вставить("EAN8", "");
	ПоследниеШтрихКоды.Вставить("EAN13", "");
	ПоследниеШтрихКоды.Вставить("Code128", "");
	ПоследниеШтрихКоды.Вставить("Авто", "");

    Если НЕ ЗначениеЗаполнено(ТекСтруктураИерархии) Тогда
		ТекСтруктураИерархии = Константы.торо_ИерархияДляВводаНовыхОР.Получить();
	КонецЕсли;  
	
	ИспользоватьКартинкиТОР = Константы.торо_ИспользоватьКартинкиТиповыхОбъектов.Получить();
	Если ИспользоватьКартинкиТОР Тогда 
		Элементы.Дерево.ПутьКДаннымКартинкиСтроки = "Дерево.КартинкаОтображаемая";
	Иначе 
		Элементы.Дерево.ПутьКДаннымКартинкиСтроки = "Дерево.Картинка";
	КонецЕсли;
	
	Элементы.ДеревоПометка.ШиринаЭлемента = 4; 
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ) 
	ОбновитьИерархиюКлиент();
КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытием(Отказ, ЗавершениеРаботы, ТекстПредупреждения, СтандартнаяОбработка)
	Если Модифицированность Тогда 
		Отказ = Истина;
		Если Не ЗавершениеРаботы Тогда 
			СтандартнаяОбработка = Ложь;
			ОписаниеОповещения = Новый ОписаниеОповещения("ПередЗакрытиемЗавершение", ЭтаФорма);
			ПоказатьВопрос(ОписаниеОповещения, 
				НСтр("ru = 'Записать изменения перед закрытием?'"), 
				РежимДиалогаВопрос.ДаНетОтмена);
		Иначе
			ТекстПредупреждения = "При закрытии формы все данные будут утеряны. Продолжить?";
		КонецЕсли;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ПередЗакрытиемЗавершение(Результат, ДопПараметры) Экспорт 
	Если Результат = КодВозвратаДиалога.Да Тогда
		ЗаписатьИЗакрыть(Неопределено);
	ИначеЕсли Результат = КодВозвратаДиалога.Нет Тогда 
		Модифицированность = Ложь;
		Закрыть();
	КонецЕсли;
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура НастройкаИерархии(Команда)
	
	Если НЕ Модифицированность Тогда 
		ВыбратьНовуюИерархию();
	Иначе
		ОписаниеОповещения = Новый ОписаниеОповещения("НастройкаИерархииВопрос", ЭтаФорма);
		ПоказатьВопрос(ОписаниеОповещения, 
			НСтр("ru = 'После выбора новой структуры иерархии не записанные штрихкоды будут очищены. Продолжить?'"), 
			РежимДиалогаВопрос.ДаНет);
	КонецЕсли;
	
КонецПроцедуры 

&НаКлиенте
Процедура УстановитьФлажкиДерево(Команда)
	
	ПостроитьДеревоОбъектовРемонта();	
	
	ОбновитьПометкиОР(Истина);
	
	УстановитьЗаголовокКнопкиСгенерировать();
		
КонецПроцедуры
	
&НаКлиенте
Процедура СнятьФлажкиДерево(Команда)
	
	ОбновитьПометкиОР(Ложь); 
	
	УстановитьЗаголовокКнопкиСгенерировать();
	
КонецПроцедуры

&НаКлиенте
Процедура ИнвертироватьФлажкиДерево(Команда) 
	
	ПостроитьДеревоОбъектовРемонта();
	
	ОбновитьПометкиОР(Неопределено); 
	ОбновитьПометкиГрупп();
	
	УстановитьЗаголовокКнопкиСгенерировать();
	
КонецПроцедуры

&НаКлиенте
Процедура Сгенерировать(Команда)
	
	Если Не ЗначениеЗаполнено(ТипШтрихКодаДляГенерации) Тогда 
		ОбщегоНазначенияКлиент.СообщитьПользователю(НСтр("ru = 'Выберите тип генерируемых штрихкодов.'"));
		Возврат;
	КонецЕсли;
	
	ПомеченныеСтроки = Новый Массив;
	ПолучитьПомеченныеСтрокиДерева(Дерево, ПомеченныеСтроки);
	
 	МассивШК = СгенерироватьНаСервере(ПомеченныеСтроки.Количество());
	
	Итерация = 0;
	Для Каждого СтрокаДерева Из ПомеченныеСтроки Цикл
		
		Если Итерация < МассивШК.Количество() Тогда 
			
			СтрокаДерева.ШтрихКод = МассивШК[Итерация];
			СтрокаДерева.ТипШтрихкода = ТипШтрихКодаДляГенерации;
			
			Итерация = Итерация + 1; 
			Модифицированность = Истина;
		Иначе
			Прервать;
		КонецЕсли;
	КонецЦикла; 
	
	ПереопределитьМаксимумыКодов(СтрокаДерева.ШтрихКод, СтрокаДерева.ТипШтрихкода);

КонецПроцедуры

&НаКлиенте
Процедура СвернутьВсе(Команда)
	Элементы.Дерево.Свернуть(Дерево.ПолучитьЭлементы()[0].ПолучитьИдентификатор());
КонецПроцедуры

&НаКлиенте
Процедура РазвернутьВсе(Команда)
	Элементы.Дерево.Развернуть(Дерево.ПолучитьЭлементы()[0].ПолучитьИдентификатор(), Истина);
КонецПроцедуры

&НаКлиенте
Процедура ЗаписатьИЗакрыть(Команда)
	ЗакрытьФорму = Истина;
	ЗаписатьНаСервере(, ЗакрытьФорму);
	Если ЗакрытьФорму Тогда 
		Модифицированность = Ложь;
		Закрыть();
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура Записать(Команда)
	
	СнятьМодифицированность = Истина; 
	
	ЗаписатьНаСервере(, СнятьМодифицированность); 
	
	ОбновитьПометкиГрупп();
	УстановитьЗаголовокКнопкиСгенерировать();
	
	Если СнятьМодифицированность Тогда 
		Модифицированность = Ложь;
	КонецЕсли; 
	
КонецПроцедуры

&НаКлиенте
Процедура Печать(Команда) 
	
	IDПомеченныхСтрок = Новый Массив;
	ПолучитьПомеченныеСтрокиДерева(Дерево, IDПомеченныхСтрок, Истина);
	
	Если IDПомеченныхСтрок.Количество() = 0 Тогда 
		ОбщегоНазначенияКлиент.СообщитьПользователю(НСтр("ru = 'Выберите объекты ремонтов для печати.'"));
		Возврат;
	КонецЕсли; 
	
	ТабличныйДокумент = Новый ТабличныйДокумент;
	ПечатьНаСервере(IDПомеченныхСтрок, ТабличныйДокумент);   
	
	ОткрытьФорму("Обработка.торо_ГенерацияШтрихКодов.Форма.ФормаПечатиШтрихКодов", Новый Структура("Макет", ТабличныйДокумент));  
	
КонецПроцедуры

&НаКлиенте
Процедура Обновить(Команда)
	Если Не Модифицированность Тогда 
		ОбновитьИерархиюКлиент();
	Иначе
		ОписаниеОповещения = Новый ОписаниеОповещения("ОбновитьВопрос", ЭтаФорма);
		ПоказатьВопрос(ОписаниеОповещения, 
			НСтр("ru = 'После обновления не записанные штрихкоды будут очищены. Продолжить?'"), 
			РежимДиалогаВопрос.ДаНет);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыДерево

&НаКлиенте
Процедура ДеревоПередРазворачиванием(Элемент, Строка, Отказ)
	
	СтрокаДерева = Дерево.НайтиПоИдентификатору(Строка);
	Если СтрокаДерева = Неопределено Или Не ЗначениеЗаполнено(СтрокаДерева.Ссылка) Тогда
		Возврат;
	КонецЕсли; 
	
	СвязиОбновлялись = СтрокаДерева.СвязиОбновлялись;
	торо_РаботаСИерархией20Клиент.ДеревоПередРазворачиванием(Дерево, Строка, СтруктураПараметровИерархии);
	
	Если Не СвязиОбновлялись Тогда
		Если ИспользоватьКартинкиТОР Тогда
			ОбновитьКартинкиЭлементовДереваНаСервере(Строка);
		КонецЕсли;  
		
		ЗаполнитьШтрихкоды(Строка);
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ДеревоШтрихКодОткрытие(Элемент, СтандартнаяОбработка)
	ТекДанные = Элементы.Дерево.ТекущиеДанные;
	Если ТекДанные <> Неопределено Тогда
		СтандартнаяОбработка = Ложь; 
		ПараметрыФормы = Новый Структура("ШтрихКод, ТипШтрихКода, РежимПросмотра, Максимумы", 
			ТекДанные.ШтрихКод, ТекДанные.ТипШтрихКода, НЕ ТекДанные.Пометка, ПоследниеШтрихКоды);
		ОписаниеОповещения = Новый ОписаниеОповещения("ДеревоШтрихКодОткрытиеЗавершение", ЭтотОбъект, ПараметрыФормы);
		ОткрытьФорму("Обработка.торо_ГенерацияШтрихКодов.Форма.ФормаШтрихкода", ПараметрыФормы, , , , , ОписаниеОповещения);
	КонецЕсли; 
КонецПроцедуры

&НаКлиенте
Процедура ДеревоШтрихКодОткрытиеЗавершение(Результат, ДопПараметры) Экспорт
	Если Результат <> Неопределено И Результат.ШтрихКод <> ДопПараметры.ШтрихКод Тогда
		ТекДанные = Элементы.Дерево.ТекущиеДанные;
		ТекДанные.ШтрихКод = Результат.ШтрихКод;
		ТекДанные.ТипШтрихкода = Результат.ТипШтрихКода; 
		ПереопределитьМаксимумыКодов(ТекДанные.ШтрихКод, ТекДанные.ТипШтрихкода);
		Модифицированность = Истина;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ДеревоНаименованиеОчистка(Элемент, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
КонецПроцедуры

&НаКлиенте
Процедура ДеревоШтрихКодОчистка(Элемент, СтандартнаяОбработка)
	СтандартнаяОбработка = Ложь;
	ТекДанные = Элементы.Дерево.ТекущиеДанные;
	Если ТекДанные <> Неопределено И ТекДанные.Пометка И Не ТекДанные.ЭтоГруппа Тогда 
		Если ТекДанные.ШтрихКод = ПоследниеШтрихКоды[торо_ОбщегоНазначенияКлиентСервер.ТипШтрихКодаСтрокой(ТекДанные.ТипШтрихКода)] Тогда 
			ПереопределитьМаксимумыКодов(ТекДанные.ШтрихКод, ТекДанные.ТипШтрихКода, Истина); 
		КонецЕсли;
		ТекДанные.ШтрихКод = "";
		ТекДанные.ТипШтрихКода = Неопределено;
		Модифицированность = Истина;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ДеревоНаименованиеОткрытие(Элемент, СтандартнаяОбработка)
	ТекДанные = Элементы.Дерево.ТекущиеДанные;
	Если ТекДанные <> Неопределено Тогда
		СтандартнаяОбработка = Ложь;
		ПоказатьЗначение(, ТекДанные.Ссылка);
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ДеревоПометкаПриИзменении(Элемент)
	
	ТекущиеДанные = Элементы.Дерево.ТекущиеДанные;
	Если ТекущиеДанные <> Неопределено Тогда
		
		ТекущиеДанные.Пометка = ?(ТекущиеДанные.Пометка = 2, 0, ТекущиеДанные.Пометка);
		
		Если ТекущиеДанные.ЕстьПодчиненные Тогда
			ПостроитьДеревоОбъектовРемонта(, ТекущиеДанные.ПолучитьИдентификатор());
			
			ПроставитьПометкуУПодчиненнныхЭлементовДерева(ТекущиеДанные, ТекущиеДанные.Пометка);
		КонецЕсли;
	
		ОбновитьПометкиГрупп();
		УстановитьЗаголовокКнопкиСгенерировать();
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура ЗаписатьНаСервере(Строка = Неопределено, ЗакрытьФорму = Истина)
	
	Если Строка = Неопределено Тогда 
		ПодчСтроки = Дерево.ПолучитьЭлементы(); 
	Иначе
		ПодчСтроки = Строка.ПолучитьЭлементы();
	КонецЕсли; 
	
	Для Каждого ПодчСтрока Из ПодчСтроки Цикл 
		Если ПодчСтрока.Пометка И Не ПодчСтрока.ЭтоГруппа Тогда 
			СпрОбъект = ПодчСтрока.Ссылка.ПолучитьОбъект();
			Если СпрОбъект.ШтрихКод <> ПодчСтрока.ШтрихКод Тогда
				СпрОбъект.ШтрихКод = ПодчСтрока.ШтрихКод;
				СпрОбъект.ТипШтрихКода = ПодчСтрока.ТипШтрихкода;
				Ошибка = Ложь;
				Попытка 
					СпрОбъект.Записать();
				Исключение
					Ошибка = Истина;
					ЗакрытьФорму = Ложь;
					ОбщегоНазначения.СообщитьПользователю(СтрШаблон(
						НСтр("ru = 'При записи объекта ремонта ""%1"" произошла ошибка'"), ПодчСтрока.Ссылка));
				КонецПопытки;
					
				Если Не Ошибка Тогда 
					ПодчСтрока.Пометка = Ложь;
				КонецЕсли;
					
			КонецЕсли;
		КонецЕсли; 
		
		ЗаписатьНаСервере(ПодчСтрока, ЗакрытьФорму);
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура ПолучитьПомеченныеСтрокиДерева(СтрокаДерева, ПомеченныеСтроки, ПолучитьIDСтрок = Ложь)
	
	Для Каждого ПодчиненнаяСтрока Из СтрокаДерева.ПолучитьЭлементы() Цикл 
		
		Если ПодчиненнаяСтрока.Пометка 
			И Не ПодчиненнаяСтрока.ЭтоГруппа Тогда 
			
			ДобавляемоеЗначение = ?(ПолучитьIDСтрок, ПодчиненнаяСтрока.ПолучитьИдентификатор(), ПодчиненнаяСтрока); 
			ПомеченныеСтроки.Добавить(ДобавляемоеЗначение);
		КонецЕсли;

		ПолучитьПомеченныеСтрокиДерева(ПодчиненнаяСтрока, ПомеченныеСтроки, ПолучитьIDСтрок);
	КонецЦикла;
	
КонецПроцедуры

&НаСервере
Функция СгенерироватьНаСервере(Количество)
	ПоследнийШК = ПоследниеШтрихКоды[торо_ОбщегоНазначенияКлиентСервер.ТипШтрихКодаСтрокой(ТипШтрихКодаДляГенерации)];
	Возврат Обработки.торо_ГенерацияШтрихКодов.СгенерироватьШтрихкодыМассово(ТипШтрихКодаДляГенерации, Количество, ПоследнийШК);
КонецФункции

&НаСервере
Процедура ОбработатьИзменениеИерархииНаСервере()
		
	торо_РаботаСИерархией20.ЗаполнитьПараметрыИерархии(ЭтаФорма); 
	СтруктураПараметровИерархии.КонечныеЭлементыВДереве = Истина;
	
	ПустойРодитель = торо_РаботаСИерархией20.ПустойРодительПоСтруктуреИерархии(СтруктураПараметровИерархии);
	торо_РаботаСИерархией20КлиентСервер.НачатьЗаполнениеДереваИерархии(Дерево, СтруктураПараметровИерархии, ПустойРодитель);
	
	КореньДерева = Дерево.ПолучитьЭлементы()[0];
	КореньДерева.ЭтоГруппа = Истина;
	КореньДерева.ЕстьПодчиненные = Истина; 
	
	Если ИспользоватьКартинкиТОР Тогда
		ОбновитьКартинкиЭлементовДереваНаСервере(КореньДерева.ПолучитьИдентификатор());
	КонецЕсли; 
	
	// СтандартныеПодсистемы.Мультиязычность
	МультиязычностьСервер.ПриСозданииНаСервере(ЭтаФорма);
	// СтандартныеПодсистемы.Мультиязычность

КонецПроцедуры

&НаСервере
Процедура ЗаполнитьШтрихкоды(ИдентификаторСтроки) 
	
	СтрокаДерева = Дерево.НайтиПоИдентификатору(ИдентификаторСтроки);
	
	ЭлементыПодчиненнойСтроки = СтрокаДерева.ПолучитьЭлементы();
	ОбъектыРемонтов = Новый Массив;
	Для Каждого ПодчиненнаяСтрока Из ЭлементыПодчиненнойСтроки Цикл 
		Если ТипЗнч(ПодчиненнаяСтрока.Ссылка) = Тип("СправочникСсылка.торо_ОбъектыРемонта") И Не ПодчиненнаяСтрока.ЭтоГруппа Тогда
			ОбъектыРемонтов.Добавить(ПодчиненнаяСтрока.Ссылка);	
		КонецЕсли;
	КонецЦикла;
	
	РеквизитыОбъектовРемонта = ОбщегоНазначения.ЗначенияРеквизитовОбъектов(ОбъектыРемонтов, "Штрихкод, ТипШтрихкода");
	
	Для Каждого ПодчиненнаяСтрока Из ЭлементыПодчиненнойСтроки Цикл 
		Если ТипЗнч(ПодчиненнаяСтрока.Ссылка) = Тип("СправочникСсылка.торо_ОбъектыРемонта") И Не ПодчиненнаяСтрока.ЭтоГруппа Тогда
			РеквизитыОбъектаРемонта = РеквизитыОбъектовРемонта.Получить(ПодчиненнаяСтрока.Ссылка);
			ЗаполнитьЗначенияСвойств(ПодчиненнаяСтрока, РеквизитыОбъектаРемонта);
			ПереопределитьМаксимумыКодов(ПодчиненнаяСтрока.ШтрихКод, ПодчиненнаяСтрока.ТипШтрихкода);
		КонецЕсли;
	КонецЦикла
	
КонецПроцедуры

&НаСервере
Процедура ПереопределитьМаксимумыКодов(ШтрихКод, Тип, ИсключитьШК = Ложь)
	ТипСтрокой = торо_ОбщегоНазначенияКлиентСервер.ТипШтрихКодаСтрокой(Тип);
	Если ИсключитьШК Тогда 
		ПоследниеШтрихКоды.Вставить(ТипСтрокой, 
			Обработки.торо_ГенерацияШтрихКодов.ПолучитьПоследнийШтрихкод(Тип, ШтрихКод));
	ИначеЕсли ЗначениеЗаполнено(ШтрихКод) Тогда
		ПоследниеШтрихКоды.Вставить(ТипСтрокой, ШтрихКод);	
	КонецЕсли;
КонецПроцедуры

&НаСервере 
Процедура ОбновитьКартинкиЭлементовДереваНаСервере(ИдентификаторСтроки)
	
	торо_РаботаСИерархией20.ОбновитьКартинкиЭлементовДереваНаСервере(Дерево, ИдентификаторСтроки);
	
КонецПроцедуры

&НаКлиенте
Процедура НастройкаИерархииЗавершение(Результат, ДопПараметры) Экспорт
	Если ЗначениеЗаполнено(Результат) И ТекСтруктураИерархии <> Результат Тогда 
		ТекСтруктураИерархии = Результат; 
		ОбновитьИерархиюКлиент();
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ОбновитьИерархиюКлиент()
	
	Дерево.ПолучитьЭлементы().Очистить(); 
	
	ОбработатьИзменениеИерархииНаСервере(); 
	
	Элементы.Дерево.Развернуть(Дерево.ПолучитьЭлементы()[0].ПолучитьИдентификатор(), Ложь);
	
	Модифицированность = Ложь;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбновитьПометкиОР(Пометка, СтрокаДерева = Неопределено)

	Если СтрокаДерева = Неопределено Тогда 
		СтрокаДерева = Дерево;
	КонецЕсли;	
	
	Для Каждого ПодчиненнаяСтрока Из СтрокаДерева.ПолучитьЭлементы() Цикл 
		
		Если Пометка = Ложь Тогда 
			ПодчиненнаяСтрока.Пометка = 0;
		ИначеЕсли Пометка = Истина Тогда
			ПодчиненнаяСтрока.Пометка = 1;
		Иначе
			ПодчиненнаяСтрока.Пометка = ?(ПодчиненнаяСтрока.Пометка = 0, 1, 0);
		КонецЕсли;		
		
		ОбновитьПометкиОР(Пометка, ПодчиненнаяСтрока);
	КонецЦикла; 
	
КонецПроцедуры

&НаСервере
Процедура ПечатьНаСервере(МассивСтрокДерева, ТабДок)
	Макет = Обработки.торо_ГенерацияШтрихКодов.ПолучитьМакет("ШтрихКодОР");
	Для Каждого ИДСтроки Из МассивСтрокДерева Цикл 
		Строка = Дерево.НайтиПоИдентификатору(ИДСтроки);
		Если ЗначениеЗаполнено(Строка.ШтрихКод) Тогда 
			
			ОбластьШапка = Макет.ПолучитьОбласть("Шапка");
			ЗаполнитьЗначенияСвойств(ОбластьШапка.Параметры, Строка);
			ТабДок.Вывести(ОбластьШапка);
			
			ОбластьШК = Макет.ПолучитьОбласть("ОбластьШтрихкод");
			Обработки.торо_ГенерацияШтрихКодов.ОтрисоватьШтрихКод(Строка.ШтрихКод, Строка.ТипШтрихКода, ОбластьШК);
			ТабДок.Вывести(ОбластьШК);
		КонецЕсли;
	КонецЦикла;
КонецПроцедуры

&НаКлиенте
Процедура ВыбратьНовуюИерархию()
	ОписаниеОповещения = Новый ОписаниеОповещения("НастройкаИерархииЗавершение", ЭтаФорма);
	
	ПараметрыОткрытия = Новый Структура;
	ПараметрыОткрытия.Вставить("ТекущаяСтрока", ТекСтруктураИерархии);
	
	ОткрытьФорму("Справочник.торо_СтруктурыОР.ФормаВыбора", ПараметрыОткрытия, 
		ЭтаФорма, ЭтаФорма.УникальныйИдентификатор,,, ОписаниеОповещения, РежимОткрытияОкнаФормы.БлокироватьОкноВладельца);
КонецПроцедуры

&НаКлиенте
Процедура НастройкаИерархииВопрос(Результат, ДопПараметры) Экспорт
	Если Результат = КодВозвратаДиалога.Да Тогда 
		ВыбратьНовуюИерархию();
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ОбновитьВопрос(Результат, ДопПараметры) Экспорт
	Если Результат = КодВозвратаДиалога.Да Тогда 
		ОбновитьИерархиюКлиент();
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ПроставитьПометкуУПодчиненнныхЭлементовДерева(СтрокаДерева, Пометка)
	
	Для Каждого ПодчиненнаяСтрока Из СтрокаДерева.ПолучитьЭлементы() Цикл
		ПодчиненнаяСтрока.Пометка = Пометка;
		ПроставитьПометкуУПодчиненнныхЭлементовДерева(ПодчиненнаяСтрока, Пометка);
	КонецЦикла;
	
КонецПроцедуры

&НаСервере
Процедура ПостроитьДеревоОбъектовРемонта(Знач СтрокаДерева = Неопределено, ИдентификаторСтроки = "")
	
	Если ЗначениеЗаполнено(ИдентификаторСтроки) Тогда
		СтрокаДерева = Дерево.НайтиПоИдентификатору(ИдентификаторСтроки);
	ИначеЕсли СтрокаДерева = Неопределено Тогда
		СтрокаДерева = Дерево;	
	КонецЕсли;	
	
	Для Каждого ПодчиненнаяСтрока Из СтрокаДерева.ПолучитьЭлементы() Цикл
				
		Если Не ПодчиненнаяСтрока.СвязиОбновлялись
			И ПодчиненнаяСтрока.ЕстьПодчиненные
			И ЗначениеЗаполнено(ПодчиненнаяСтрока.Ссылка) Тогда
			
			МассивРодителей = Новый Массив;
			МассивРодителей.Добавить(ПодчиненнаяСтрока.Ссылка);
			
			ПодчиненнаяСтрока.ПолучитьЭлементы().Очистить();
			
			МассивСтрок = торо_РаботаСИерархией20.ПолучитьМассивСтрокОдногоУровняИерархииСКоличествомПодчиненных(СтруктураПараметровИерархии, МассивРодителей,,, Истина);
			торо_РаботаСИерархией20КлиентСервер.ДобавитьСтрокиВДеревоСФиктивнымиПодчиненнымиЭлементами(МассивСтрок, ПодчиненнаяСтрока);
			
			ПодчиненнаяСтрока.СвязиОбновлялись = Истина;
			
			ИдентификаторСтроки = ПодчиненнаяСтрока.ПолучитьИдентификатор();
			
			Если ИспользоватьКартинкиТОР Тогда
				ОбновитьКартинкиЭлементовДереваНаСервере(ИдентификаторСтроки);
			КонецЕсли;
			ЗаполнитьШтрихкоды(ИдентификаторСтроки);
		КонецЕсли;	
		
		ПостроитьДеревоОбъектовРемонта(ПодчиненнаяСтрока);	
	КонецЦикла;
	
КонецПроцедуры	

&НаКлиенте
Процедура ОбновитьПометкиГрупп()
	
	Для Каждого ПодчиненнаяСтрока Из Дерево.ПолучитьЭлементы() Цикл
		ПодчиненнаяСтрока.Пометка = ЗначениеПометкиПоПодчиненнымСтрокам(ПодчиненнаяСтрока);
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Функция ЗначениеПометкиПоПодчиненнымСтрокам(СтрокаДерева)
	
	Пометки = Новый Массив;
	
	Для Каждого ПодчиненнаяСтрока Из СтрокаДерева.ПолучитьЭлементы() Цикл
		
		Если ПодчиненнаяСтрока.ПолучитьЭлементы().Количество() > 0 Тогда
			ПодчиненнаяСтрока.Пометка = ЗначениеПометкиПоПодчиненнымСтрокам(ПодчиненнаяСтрока);	
		КонецЕсли;
		
		Если Пометки.Найти(ПодчиненнаяСтрока.Пометка) = Неопределено Тогда
			Пометки.Добавить(ПодчиненнаяСтрока.Пометка);
		КонецЕсли;	
	КонецЦикла;
	
	Если Пометки.Количество() = 2 Тогда
		Результат = 2;
	ИначеЕсли Пометки.Количество() = 1 Тогда
		Результат = Пометки[0];
	Иначе
		Результат = СтрокаДерева.Пометка;
	КонецЕсли;	
	
	Возврат Результат;
	
КонецФункции

&НаКлиенте
Процедура УстановитьЗаголовокКнопкиСгенерировать()
	
	Элементы.Сгенерировать.Заголовок = СтрШаблон("Сгенерировать (%1)", КоличествоПомеченныхОбъектовРемонта());	
	
КонецПроцедуры	

&НаКлиенте
Функция КоличествоПомеченныхОбъектовРемонта(СтрокаДерева = Неопределено, Количество = 0) 
	
	Если СтрокаДерева = Неопределено Тогда
		СтрокаДерева = Дерево;
	КонецЕсли;	
	
	Для Каждого ПодчиненнаяСтрока Из СтрокаДерева.ПолучитьЭлементы() Цикл
		
		ЕстьПодчиненные = ПодчиненнаяСтрока.ПолучитьЭлементы().Количество() > 0;
		Если ПодчиненнаяСтрока.Пометка = 1 И Не ПодчиненнаяСтрока.ЭтоГруппа И Не ЕстьПодчиненные Тогда	
			Количество = Количество + 1;
		ИначеЕсли ПодчиненнаяСтрока.Пометка = 1 И Не ПодчиненнаяСтрока.ЭтоГруппа И ЕстьПодчиненные Тогда
			Количество = Количество + 1 + КоличествоПомеченныхОбъектовРемонта(ПодчиненнаяСтрока);	
		Иначе	
			Количество = Количество + КоличествоПомеченныхОбъектовРемонта(ПодчиненнаяСтрока);
		КонецЕсли;
	КонецЦикла;	
	
	Возврат Количество;
	
КонецФункции

#КонецОбласти
