#Если НЕ ВебКлиент Тогда

#Область ПрограммныйИнтерфейс

// Инициирует передачу документа в Project.
//
// Параметры:
//  ПараметрКоманды				 - ДокументСсылка.торо_ЗаявкаНаРемонт, Массив - параметр команды выгрузки в Project.
//  ПараметрыВыполненияКоманды - Структура - структура параметров выполнения команды.
//  ПереданМассивID				 - Булево - если Истина, то в массиве лежат ID ремонтов, иначе - ссылки на документы.
//
Процедура ОтправитьВProject(ПараметрКоманды, ПараметрыВыполненияКоманды, ПереданМассивID = Ложь) Экспорт
	
	ДопПараметры = Новый Структура("ПараметрКоманды,ПереданМассивID", ПараметрКоманды, ПереданМассивID);
	
	ОписаниеОповещения = новый ОписаниеОповещения("ОбработкаОтветаПользователя", ЭтотОбъект, ДопПараметры);
	
	Кнопки = Новый СписокЗначений;
	Кнопки.Добавить(1, НСтр("ru = 'В MS Project напрямую'"));
	Кнопки.Добавить(2, НСтр("ru = 'В промежуточный XML-файл'"));
	Кнопки.Добавить(3, НСтр("ru = 'Отмена'"));
	
	ПоказатьВопрос(ОписаниеОповещения, НСтр("ru = 'Укажите желаемый способ передачи данных:'"), Кнопки,, 1, НСтр("ru = 'Интеграция с MS Project'"));
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Процедура ОбработкаОтветаПользователя(ВариантВыгрузки, ДополнительныеПараметры) Экспорт
	
	Если ВариантВыгрузки = 0 Тогда
		Возврат;
	КонецЕсли;
	
	ДополнительныеПараметры.Вставить("ВариантВыгрузки", ВариантВыгрузки);
	
	ОписаниеОповещения = Новый ОписаниеОповещения("ВыгрузитьПослеВыбораФайла", ЭтотОбъект, ДополнительныеПараметры);
		
	Если ВариантВыгрузки = 1 Тогда
		ОткрытьДиалогВыбораФайла(ОписаниеОповещения, "mpp");
	ИначеЕсли ВариантВыгрузки = 2 Тогда
		ОткрытьДиалогВыбораФайла(ОписаниеОповещения, "xml");
	КонецЕсли;
	
КонецПроцедуры

Процедура ВыгрузитьПослеВыбораФайла(ВыбранныеФайлы, ДополнительныеПараметры) Экспорт
	
	Если ВыбранныеФайлы = Неопределено Тогда
		Возврат;
	КонецЕсли;
	
	Если ВыбранныеФайлы.Количество() = 0 Тогда
		Возврат;
	КонецЕсли;
	
	ПутьКФайлу = ВыбранныеФайлы[0];
	
	ВариантВыгрузки = ДополнительныеПараметры.ВариантВыгрузки;
	ПараметрКоманды = ДополнительныеПараметры.ПараметрКоманды;
	ПереданМассивID = ДополнительныеПараметры.ПереданМассивID;
	
	Если ВариантВыгрузки = 1 Тогда
		Если ТипЗнч(ПараметрКоманды) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") Тогда
			ВыгрузитьЗаявку(ПутьКФайлу, ПараметрКоманды);
		ИначеЕсли ТипЗнч(ПараметрКоманды) = Тип("Массив") Тогда
			ВыгрузитьЗаявки(ПутьКФайлу, ПараметрКоманды, Не ПереданМассивID);
		КонецЕсли;
	ИначеЕсли ВариантВыгрузки = 2 Тогда
		Если ТипЗнч(ПараметрКоманды) = Тип("ДокументСсылка.торо_ЗаявкаНаРемонт") Тогда
			ВыгрузитьЗаявкуXML(ПутьКФайлу, ПараметрКоманды);
		ИначеЕсли ТипЗнч(ПараметрКоманды) = Тип("Массив") Тогда
			ВыгрузитьЗаявкиXML(ПутьКФайлу, ПараметрКоманды, Не ПереданМассивID);
		КонецЕсли; 
	КонецЕсли;
	
КонецПроцедуры	
	
#Область В_PROJECT_ПО_COM

Процедура ВыгрузитьЗаявку(ПутьКФайлу, Заявка) 
	
	Если ЗначениеЗаполнено(ПутьКФайлу) Тогда 
		
		Ссылки = Новый Массив;
		Ссылки.Добавить(Заявка);
		ПараметрыСоответствия = Новый Структура("Ссылки", Ссылки);
		Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьСоответствие(ПараметрыСоответствия);
		
		Если ПередатьДанныеПоCOM(Данные, ПутьКФайлу) Тогда
			ПоказатьПредупреждение(, НСтр("ru = 'Выгрузка в MS Project завершена.'"), 5, НСтр("ru = 'Интеграция с MS Project'"));
		КонецЕсли;
		
	КонецЕсли;

КонецПроцедуры

Процедура ВыгрузитьЗаявки(ПутьКФайлу, МассивЗаявок, ПоискПоСсылке = ИСТИНА) 

	Если ЗначениеЗаполнено(ПутьКФайлу) Тогда 
		ПараметрыСоответствия = Новый Структура("Ссылки", МассивЗаявок);
		Если МассивЗаявок.Количество()>0 И ТипЗнч(МассивЗаявок[0]) = Тип("ДокументСсылка.торо_ОстановочныеРемонты") Тогда
			Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьОстСоответствие(ПараметрыСоответствия);
		ИначеЕсли МассивЗаявок.Количество()>0 И ТипЗнч(МассивЗаявок[0]) = Тип("Структура") Тогда
			Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьСоответствиеПланРабот(ПараметрыСоответствия);
		Иначе
			Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьСоответствие(ПараметрыСоответствия, ПоискПоСсылке);
		КонецЕсли;
		Если ПередатьДанныеПоCOM(Данные, ПутьКФайлу) Тогда
			ПоказатьПредупреждение(, НСтр("ru = 'Выгрузка в MS Project завершена.'"), 5, НСтр("ru = 'Интеграция с MS Project'"));
		КонецЕсли;
	КонецЕсли;
	
КонецПроцедуры

Функция ПередатьДанныеПоCOM(Данные, ПутьКФайлу)
	
	Если ТипЗнч(Данные) <> Тип("Соответствие") Тогда
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Данные неверны.'"));
		Возврат Ложь;
	КонецЕсли;
	
	#Если МобильныйКлиент Тогда
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Функционал недоступен в мобильном клиенте.'"));
		Возврат Ложь;
	#Иначе	
	// Открываем связь с приложением
	Попытка
		prjApp = Новый COMОбъект("MSProject.Application");
	Исключение
	    ПоказатьПредупреждение(Неопределено, Нстр("ru = 'Ошибка подключения к MS Project, возможно MS Project не установлен на данном компьютере'"),,"Ошибка");
		Возврат Ложь;
	КонецПопытки;   
	
	Попытка
		Если prjApp.Projects.Count() > 0 Тогда
			prjApp = Неопределено;
			ПоказатьПредупреждение(Неопределено,НСтр("ru = 'Приложение Microsoft Project уже запущено. Загрузка или выгрузка не может быть произведена'"));
			Возврат Ложь;
		КонецЕсли;
	Исключение
		// Если в проджекте открыто модальное окно, то вываливается ошибка "Поле объекта не обнаружено (Projects)".
		ПоказатьПредупреждение(Неопределено,НСтр("ru = 'Приложение Microsoft Project уже запущено. Загрузка или выгрузка не может быть произведена'"));
		Возврат Ложь;
	КонецПопытки;
	
	Попытка
		
		Project = prjApp.Projects.Add();
		
		СвойстваПроекта = Данные.Получить("СвойстваПроекта");
		Если СвойстваПроекта <> Неопределено Тогда
			Если СвойстваПроекта.Свойство("ИДЗадач") Тогда
				// Project.ProjectNotes = СвойстваПроекта.ИДЗадач;
				ИДСвойства = 188744016; // Text30
				prjApp.CustomFieldRename(ИДСвойства, "TasksID"); 
				prjApp.CustomFieldValueList(ИДСвойства, Ложь, "", Истина, Ложь);
				prjApp.CustomFieldPropertiesEx(ИДСвойства, 2);
				Для каждого ИДЗадачи из СвойстваПроекта.ИДЗадач Цикл 
					prjApp.CustomFieldValueListAdd(ИДСвойства, ИДЗадачи);
				КонецЦикла;
			КонецЕсли;
		КонецЕсли;
		
		ствРесурсИД = Новый Соответствие;
		Ресурсы = Данные.Получить("Ресурсы");
		Если Ресурсы <> Неопределено Тогда
			
			Для каждого кзРесурс из Ресурсы Цикл
				Свойства = кзРесурс.Значение;
				
				Resource = Project.Resources.Add(Свойства.Name);
							 
				ЗаполнитьЗначенияСвойств(Resource, Свойства,, "CanLevel");
				
				ствРесурсИД.Вставить(кзРесурс.Ключ, Новый Структура("ID, Type",Resource.ID, Resource.Type));
			КонецЦикла;
			
		КонецЕсли;
		
		нЗадачи = Данные.Получить("Задачи");
		Если нЗадачи <> Неопределено Тогда
						
			Для каждого кзЗадача из нЗадачи Цикл
				
				Свойства = кзЗадача.Значение;
				
				Task = Project.Tasks.Add(Свойства.Name);
				
				ИсточникДанных = "";
				Свойства.Свойство("ИсточникДанных", ИсточникДанных);
				Task.Text29 = ИсточникДанных;
				
				ЗаполнитьЗначенияСвойств(Task, Свойства,,"Duration, Work");
								
				Назначения = Свойства.Назначения;
				Если Назначения <> Неопределено Тогда
					
					Для каждого кзНазначение из Назначения Цикл
						
						СвойстваНазначения = кзНазначение;						
						СтруктураДанных = ствРесурсИД.Получить(СвойстваНазначения.Ресурс);
						
						Если СтруктураДанных <> Неопределено Тогда
							
							Assignment = Task.Assignments.Add(,СтруктураДанных.ID);
							ЗаполнитьЗначенияСвойств(Assignment, СвойстваНазначения);
							
						КонецЕсли;
					КонецЦикла;
					
				КонецЕсли;
				
			КонецЦикла;
			
			
			
		КонецЕсли;
		
	Исключение
		ОбщегоНазначенияКлиент.СообщитьПользователю(ОписаниеОшибки());
		prjApp.Quit();
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Ошибка заполнения файла проекта'"));
		Возврат Ложь;
	КонецПопытки;
	
	// Сохраняем в файл
	
	Project.SaveAs(ПутьКФайлу); 

	// Закрываем связь с приложением. Обязательно.
	prjApp.Quit();
	
	Возврат Истина;
	
	#КонецЕсли
	
КонецФункции

#КонецОбласти

#Область ИЗ_PROJECT_ПО_COM

Функция ЗагрузитьЗаявки(ПутьКФайлуДляЗагрузки, ТипЗагрузки = "ЗаявкаНаПлатеж") Экспорт
	
	ОбновленныеДокументы = Новый Массив;
	
	Данные = ПолучитьДанныеПоCOM(ПутьКФайлуДляЗагрузки);
	Если Данные = Неопределено Тогда
		Возврат ОбновленныеДокументы;
	КонецЕсли; 
		
	Если ТипЗагрузки = "ОстановочныйРемонт" Тогда
		Ошибки = торо_ИнтеграцияСmsProjectСервер.ОбработатьСоответствиеСДаннымиОст(Данные, ОбновленныеДокументы);
	ИначеЕсли ТипЗагрузки = "ПланРабот" Тогда
		Ошибки = торо_ИнтеграцияСmsProjectСервер.ОбработатьСоответствиеСДаннымиПланРабот(Данные, ОбновленныеДокументы);
	Иначе  // "ЗаявкаНаПлатеж"
		Ошибки = торо_ИнтеграцияСmsProjectСервер.ОбработатьСоответствиеСДанными(Данные, ОбновленныеДокументы);
	КонецЕсли;
	
	Если Не Ошибки Тогда
		ТекстПредупреждения = НСтр("ru = 'Загрузка из MS Project завершена.'");
	Иначе
		ТекстПредупреждения = НСтр("ru = 'Загрузка из MS Project произведена с ошибками.'");
	КонецЕсли;
	ПоказатьПредупреждение(, ТекстПредупреждения, 5, НСтр("ru = 'Интеграция с MS Project'"));
	
	Возврат ОбновленныеДокументы;
	
КонецФункции

Функция ПолучитьДанныеПоCOM(ПутьКФайлуДляЗагрузки)
	
	Если ПутьКФайлуДляЗагрузки = "" Тогда
		Возврат Неопределено; 
	КонецЕсли;
	
	#Если МобильныйКлиент Тогда
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Функционал недоступен в мобильном клиенте.'"));
		Возврат Неопределено;
	#Иначе	
	
	ФайлMPP = Новый Файл(ПутьКФайлуДляЗагрузки);
	ИмяФайла = ФайлMPP.Имя;
	
	// Открываем связь с приложением
	Попытка
		prjApp = Новый COMОбъект("MSProject.Application");
	Исключение
	    ПоказатьПредупреждение(Неопределено, Нстр("ru = 'Ошибка подключения к MS Project, возможно MS Project не установлен на данном компьютере'"),,НСтр("ru = 'Ошибка'"));
		Возврат Неопределено;
	КонецПопытки;  
	
	Если prjApp.Projects.Count() > 0 Тогда
		prjApp = Неопределено;
		ПоказатьПредупреждение(Неопределено,НСтр("ru = 'Приложение Microsoft Project уже запущено. Загрузка или выгрузка не может быть произведена'"));
		Возврат Неопределено;
	КонецЕсли;

	СоответствиеЗадач = Новый Соответствие;
	СоответствиеРесурсов = Новый Соответствие;
	СоответствиеНазначений = Новый Соответствие;
	
	Попытка
		// Открываем файл на диске
		prjApp.FileOpen(ПутьКФайлуДляЗагрузки);
		Для Каждого сProject из prjApp.Projects Цикл
			Если сProject.Name = ИмяФайла ИЛИ сProject.Name +".mpp" = ИмяФайла Тогда
				НашПроект = сProject;
				Прервать;
			КонецЕсли;
		КонецЦикла;
		
		
		Если НашПроект <> Неопределено Тогда
			
			ИДЗадач = Новый массив;
			Индекс = 1;
			Пока Истина Цикл
				Попытка
					ИДЗадачи = prjApp.CustomFieldValueListGetItem(188744016, 0, Индекс);
					ИДЗадач.Добавить(ИДЗадачи);
					Индекс = Индекс + 1;
				Исключение
					Прервать;
				КонецПопытки;
			КонецЦикла;
			
			Для каждого Task из НашПроект.Tasks Цикл
				ДанныеЗадачи = Новый Структура(торо_ИнтеграцияСmsProjectКлиентСервер.СтрокаСвойствЗадач());
				ID = Task.Text30;
				ЗаполнитьЗначенияСвойств(ДанныеЗадачи, Task);
				СоответствиеЗадач.Вставить(ID, ДанныеЗадачи);
				
				Индекс = ИДЗадач.Найти(ID);
				Если Индекс <> Неопределено Тогда
					ИДЗадач.Удалить(Индекс);
				КонецЕсли;
			КонецЦикла;
			
			Для каждого Resource из НашПроект.Resources Цикл
				ДанныеРесурса = Новый Структура(торо_ИнтеграцияСmsProjectКлиентСервер.СтрокаСвойствРесурсов());
				ID = Resource.Text30;
				ЗаполнитьЗначенияСвойств(ДанныеРесурса, Resource);
				СоответствиеРесурсов.Вставить(ID, ДанныеРесурса);
			КонецЦикла;
			
			СообщитьОбУдаленныхРемонтах(ИДЗадач);
		КонецЕсли;
		
		prjApp.Quit();
		
	Исключение
		ОбщегоНазначенияКлиент.СообщитьПользователю(ОписаниеОшибки());
		prjApp.Quit();
		ПоказатьПредупреждение(Неопределено, НСтр("ru = 'Ошибка при чтении файла проекта'"));
	КонецПопытки;
	
	Данные = Новый Соответствие();
	Данные.Вставить("Задачи", СоответствиеЗадач);
	Данные.Вставить("Ресурсы", СоответствиеРесурсов);
	Данные.Вставить("Назначения", СоответствиеНазначений);
	Возврат Данные;
	
	#КонецЕсли
	
КонецФункции

#КонецОбласти

#Область В_PROJECT_ЧЕРЕЗ_XML

Процедура ВыгрузитьЗаявкуXML(ПутьКФайлу, Заявка) 
	
	Если ПутьКФайлу = "" Тогда
		Возврат;
	КонецЕсли;
	
	Ссылки = Новый Массив;
	Ссылки.Добавить(Заявка);
	ПараметрыСоответствия = Новый Структура("Ссылки", Ссылки);
	
	Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьСоответствие(ПараметрыСоответствия);
	Текст = торо_ИнтеграцияСmsProjectСервер.ТекстФайлаXML(Данные);
	СохранитьТекстовыйФайлНаДиске(ПутьКФайлу, Текст);
	
	ПоказатьПредупреждение(, НСтр("ru = 'Выгрузка в MS Project завершена.'"), 5, НСтр("ru = 'Интеграция с MS Project'"));
	
КонецПроцедуры

Процедура ВыгрузитьЗаявкиXML(ПутьКФайлу, МассивЗаявок, ПоискПоСсылке = ИСТИНА) 
	
	Если ПутьКФайлу = "" Тогда
		Возврат;
	КонецЕсли;
	
	ПараметрыСоответствия = Новый Структура("Ссылки", МассивЗаявок);
	Если МассивЗаявок.Количество()>0 И ТипЗнч(МассивЗаявок[0]) = Тип("ДокументСсылка.торо_ОстановочныеРемонты") Тогда
		Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьОстСоответствие(ПараметрыСоответствия);
	ИначеЕсли МассивЗаявок.Количество()>0 И ТипЗнч(МассивЗаявок[0]) = Тип("Структура") Тогда
		Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьСоответствиеПланРабот(ПараметрыСоответствия);
	Иначе
		Данные = торо_ИнтеграцияСmsProjectСервер.ПодготовитьСоответствие(ПараметрыСоответствия, ПоискПоСсылке);
	КонецЕсли;
	Текст = торо_ИнтеграцияСmsProjectСервер.ТекстФайлаXML(Данные);
		
	СохранитьТекстовыйФайлНаДиске(ПутьКФайлу, Текст);
	
	ПоказатьПредупреждение(, НСтр("ru = 'Выгрузка в MS Project завершена.'"), 5, НСтр("ru = 'Интеграция с MS Project'"));
	
КонецПроцедуры

Процедура СохранитьТекстовыйФайлНаДиске(Путь, Текст)
	Документ = ПолучитьИзВременногоХранилища(Текст);
	Документ.Записать(Путь);
КонецПроцедуры

Процедура ОткрытьДиалогВыбораФайла(ОписаниеОповещения, Расширение = "xml")
	
	ЗаголовокДиалога = НСтр("ru = 'Выберите место сохранения файла:'");
	
	ДиалогОткрытияФайла = Новый ДиалогВыбораФайла(РежимДиалогаВыбораФайла.Сохранение);
	ДиалогОткрытияФайла.МножественныйВыбор = Ложь;
	ДиалогОткрытияФайла.Заголовок = ЗаголовокДиалога;
	
	Если Расширение = "xml" Тогда
		ДиалогОткрытияФайла.Фильтр = НСтр("ru = 'Формат XML(*.xml)|*.xml'");
	Иначе
		ДиалогОткрытияФайла.Фильтр = НСтр("ru = 'Проект(*.mpp)|*.mpp'");
	КонецЕсли;
	
	ДиалогОткрытияФайла.Показать(ОписаниеОповещения);
	
КонецПроцедуры

#КонецОбласти

#Область ИЗ_PROJECT_ЧЕРЕЗ_XML

Функция ЗагрузитьЗаявкиИзXML(Знач ПутьКФайлуДляЗагрузки, ТипЗагрузки = "ЗаявкаНаПлатеж") Экспорт
	
	Данные = ПолучитьДанныеИзXML(ПутьКФайлуДляЗагрузки);
	
	ОбновленныеДокументы = Новый Массив;
	
	Если Данные <> Неопределено Тогда
		Если ТипЗагрузки = "ОстановочныйРемонт" Тогда
			Ошибки = торо_ИнтеграцияСmsProjectСервер.ОбработатьСоответствиеСДаннымиОст(Данные, ОбновленныеДокументы);
		ИначеЕсли ТипЗагрузки = "ПланРабот" Тогда
			Ошибки = торо_ИнтеграцияСmsProjectСервер.ОбработатьСоответствиеСДаннымиПланРабот(Данные, ОбновленныеДокументы);
		Иначе  // "ЗаявкаНаПлатеж"
			Ошибки = торо_ИнтеграцияСmsProjectСервер.ОбработатьСоответствиеСДанными(Данные, ОбновленныеДокументы);
		КонецЕсли;
		
		Если Не Ошибки Тогда
			ТекстПредупреждения = НСтр("ru = 'Загрузка из MS Project завершена.'");
		Иначе
			ТекстПредупреждения = НСтр("ru = 'Загрузка из MS Project произведена с ошибками.'");
		КонецЕсли;
	Иначе
		ТекстПредупреждения = НСтр("ru = 'Ошибка при создании COM-объекта «MSXML2.DOMDocument».'");
	КонецЕсли;
	
	ПоказатьПредупреждение(, ТекстПредупреждения, 5, НСтр("ru = 'Интеграция с MS Project'"));
	
	Возврат ОбновленныеДокументы;
	
КонецФункции

Функция ПолучитьФабрикуИзМакета(Знач ИмяМакета)
	
	ТекстСхемы = торо_ИнтеграцияСmsProjectСервер.ОбщийТекстовыйМакет(ИмяМакета);
	
	ИмяВременногоФайла = ПолучитьИмяВременногоФайла("xsd");
	ТекстСхемы.Записать(ИмяВременногоФайла);
	
	Пути = Новый Массив();
	Пути.Добавить(ИмяВременногоФайла);
	
	Попытка
		Фабрика = СоздатьФабрикуXDTO(Пути);
	Исключение
		ОбщегоНазначенияКлиент.СообщитьПользователю(ОписаниеОшибки());
		Возврат Неопределено;
	КонецПопытки;
	
	Возврат Фабрика;
	
КонецФункции


Функция ПолучитьДанныеИзXML(Знач ПутьКФайлуДляЗагрузки)
	
	Данные = Неопределено;
	
	СоответствиеЗадач      = Новый Соответствие;
	СоответствиеРесурсов   = Новый Соответствие;
	СоответствиеНазначений = Новый Соответствие;
	
	ЧтениеXML = Новый ЧтениеXML;
	ЧтениеXML.ОткрытьФайл(ПутьКФайлуДляЗагрузки);
	
	Фабрика = ПолучитьФабрикуИзМакета("ProjectXMLSchema");
	ДанныеXDTO = Фабрика.ПрочитатьXML(ЧтениеXML);
	
	Если ДанныеXDTO.Tasks.Свойства().Получить("Task") = Неопределено Тогда
		Данные = Новый Соответствие();
		Данные.Вставить("НетЗадач", Истина);
		Возврат Данные;
	КонецЕсли; 
	
	ИДЗадач = Новый Массив;
	Если ДанныеXDTO.Свойства().Получить("OutlineCodes") <> Неопределено
		И ДанныеXDTO.OutlineCodes.Свойства().Получить("OutlineCode") <> Неопределено 
		И ДанныеXDTO.Свойства().Получить("ExtendedAttributes") <> Неопределено
		И ДанныеXDTO.ExtendedAttributes.Свойства().Получить("ExtendedAttribute") <> Неопределено Тогда
		
		ДопСвойствоText30 = ПолучитьКонкретныйОбъектXDTOИзСписка(ДанныеXDTO.ExtendedAttributes.ExtendedAttribute, "FieldID", "188744016");// text30
		Если ДопСвойствоText30 <> Неопределено Тогда
			ИдентификаторСпискаИДЗадач = ДопСвойствоText30.Ltuid;	
			Если ЗначениеЗаполнено(ИдентификаторСпискаИДЗадач) Тогда
				OutlineCode = ПолучитьКонкретныйОбъектXDTOИзСписка(ДанныеXDTO.OutlineCodes.OutlineCode, "Guid", ИдентификаторСпискаИДЗадач);
				Если OutlineCode <> Неопределено 
					И OutlineCode.Свойства().Получить("Values") <> Неопределено
					И OutlineCode.Values.Свойства().Получить("Value") <> Неопределено	Тогда
					Если ТипЗнч(OutlineCode.Values.Value) = Тип("СписокXDTO") Тогда
						Для каждого ЭлементСписка из OutlineCode.Values.Value Цикл
							ИДЗадач.Добавить(ЭлементСписка.Value);
						КонецЦикла;
					Иначе
					    ИДЗадач.Добавить(OutlineCode.Values.Value.Value);
					КонецЕсли;
				КонецЕсли;
			КонецЕсли;
		КонецЕсли;
		
	КонецЕсли;
	
	Если ТипЗнч(ДанныеXDTO.Tasks.Task) = Тип("СписокXDTO") Тогда
		
		Для каждого curTask из ДанныеXDTO.Tasks.Task Цикл
			
			Если curTask.Summary <> "0" Тогда // пропускаем суммарные задачи
				Продолжить;
			КонецЕсли;
			
			ДопСвойствоText30 = ПолучитьКонкретныйОбъектXDTOИзСписка(curTask.ExtendedAttribute, "FieldID", "188744016");
			IDработы = ?(ДопСвойствоText30 <> Неопределено, ДопСвойствоText30.Value, "");
			
			ДопСвойствоText29 = ПолучитьКонкретныйОбъектXDTOИзСписка(curTask.ExtendedAttribute, "FieldID", "188744015");
			ИсточникДанных = ?(ДопСвойствоText29 <> Неопределено, ДопСвойствоText29.Value, "");
			
			Если IDработы <> "" Тогда
				ДанныеЗадачи = Новый Структура(торо_ИнтеграцияСmsProjectКлиентСервер.СтрокаСвойствЗадач());
				ЗаполнитьЗначенияСвойств(ДанныеЗадачи, curTask);
				ДанныеЗадачи.Start = ПолучитьДату(ДанныеЗадачи.Start);
				ДанныеЗадачи.Finish = ПолучитьДату(ДанныеЗадачи.Finish);
				ДанныеЗадачи.Text30 = IDработы;
				ДанныеЗадачи.Text29 = ИсточникДанных;
				СоответствиеЗадач.Вставить(IDработы, ДанныеЗадачи);
				
				Индекс = ИДЗадач.Найти(IDработы);
				Если Индекс <> Неопределено Тогда
					ИДЗадач.Удалить(Индекс);
				КонецЕсли;
			КонецЕсли;
			
		КонецЦикла;
		
	КонецЕсли;
	
	СообщитьОбУдаленныхРемонтах(ИДЗадач);
	
	Данные = Новый Соответствие();
	
	Данные.Вставить("Задачи", СоответствиеЗадач);
	Данные.Вставить("Ресурсы", СоответствиеРесурсов);
	Данные.Вставить("Назначения", СоответствиеНазначений);
	
	Возврат Данные;
	
КонецФункции

// Функция - Получить дату из строкового представления типа «2014-01-06T00:00:00».
//
// Параметры:
//  Значение - Строка - Строковое представление даты.
// Возвращаемое значение: 
//		Дата - полученная дата.
Функция ПолучитьДату(Знач Значение)
	
	Если ТипЗнч(Значение) = Тип("Дата") Тогда
		Возврат Значение;
	КонецЕсли;
	
	Если Значение = "" ИЛИ Значение = Неопределено ИЛИ Значение = NULL Тогда
		Возврат Дата(1,1,1);
	КонецЕсли;
	
	// "2014-01-06T00:00:00"
	
	Год =       Лев(Значение, 4);
	Месяц =    Сред(Значение, 6,2);
	День =     Сред(Значение, 9,2);
	Час =      Сред(Значение,12,2);
	Минута =   Сред(Значение,15,2);
	Секунда =  Сред(Значение,18,2);
	
	Попытка
		Возврат Дата(Год, Месяц, День, Час, Минута, Секунда) ;
	Исключение
		Возврат Дата(1,1,1);
	КонецПопытки;
	
КонецФункции


// Функция - Получить конкретный объект XDTO из списка.
//
// Параметры:
//  ОбъектXDTO	 - ОбъектXDTO или СписокXDTO - Проверяемый объект или список объектов XDTO
//  ИмяСвойства		 - Строка - имя проверяемого свойства
//  ЗначениеСвойства - Строка - искомое занчение проверяемого свойства.
// Возвращаемое значение:
//		ОбъектXDTO - подходящий по условию фильтра, Неопределено, если таковой не найден.
Функция ПолучитьКонкретныйОбъектXDTOИзСписка(Знач ОбъектXDTO, Знач ИмяСвойства, Знач ЗначениеСвойства)
	
	ИскомыйОбъект = Неопределено;
	
	Если ТипЗнч(ОбъектXDTO) = Тип("ОбъектXDTO") Тогда
		
		Если ЕстьСвойствоXDTO(ОбъектXDTO, ИмяСвойства) И ОбъектXDTO[ИмяСвойства] = ЗначениеСвойства Тогда
			ИскомыйОбъект = ОбъектXDTO;
		КонецЕсли;
		
	ИначеЕсли ТипЗнч(ОбъектXDTO) = Тип("СписокXDTO") Тогда
		
		ПерваяИтерация = ИСТИНА;
		Для каждого ПодобъектXDTO из ОбъектXDTO Цикл
			
			Если ТипЗнч(ПодобъектXDTO) = Тип("ЗначениеXDTO") Тогда
				Продолжить;
			КонецЕсли;
			
			Если ПерваяИтерация Тогда
				Если НЕ ЕстьСвойствоXDTO(ПодобъектXDTO, ИмяСвойства) Тогда
					Прервать;
				КонецЕсли;
				ПерваяИтерация = ЛОЖЬ;
			КонецЕсли;
			
			Если ПодобъектXDTO[ИмяСвойства] = ЗначениеСвойства Тогда
				ИскомыйОбъект = ПодобъектXDTO;
				Прервать;
			КонецЕсли;
			
		КонецЦикла;
		
	КонецЕсли;
	
	Возврат ИскомыйОбъект;
КонецФункции

// Функция - Есть свойство XDTO.
//
// Параметры:
//  ОбъектXDTO		 - ОбъектXDTO - исследуемый Объект XDTO
//  ИмяСвойства		 - Строка - имя искомого свойства
//  ЗначениеСвойства - Произвольный - значение свойства (возвращается).
// Возвращаемое значение: 
//		Булево - признак наличия свойства у исследуемого объекта.
Функция ЕстьСвойствоXDTO(Знач ОбъектXDTO, Знач ИмяСвойства, ЗначениеСвойства = Неопределено)
	
	ЕстьСвойство = Неопределено;
	
	Если ТипЗнч(ОбъектXDTO) = Тип("ОбъектXDTO") Тогда
		
		ЕстьСвойство = ЛОЖЬ;
		ВсеСвойстваОбъектаXDTO = ОбъектXDTO.Свойства();
		Для каждого свойствоXDTO из ВсеСвойстваОбъектаXDTO Цикл
			Если свойствоXDTO.Имя = ИмяСвойства Тогда
				ЕстьСвойство = ИСТИНА;
				ЗначениеСвойства = ОбъектXDTO[ИмяСвойства];
				Прервать;
			КонецЕсли;
		КонецЦикла;
		
	КонецЕсли;
	
	Возврат ЕстьСвойство;
	
КонецФункции

#КонецОбласти

#Область ПОЛУЧЕНИЕ_СВОЙСТВ_COM_ОБЪЕКТОВ_PROJECT

Процедура ДобавитьСвойствоВСтруктуру(Объект, ИмяСвойства, СтруктураСвойств) Экспорт
	
	Попытка
		СтруктураСвойств.Вставить(ИмяСвойства, Объект[ИмяСвойства]);
	Исключение
		СтруктураСвойств.Вставить(ИмяСвойства, "Не удалось прочитать значение свойства «"+ИмяСвойства+"»");
	КонецПопытки;
	
КонецПроцедуры

Функция СвойстваПроекта(Проект)Экспорт
	СвойстваПроекта = Новый Структура();
	
	//    True if new or changed data relating to an external task is automatically accepted
	// when the project is opened.
	// Read/write. Boolean.
	// ДобавитьСвойствоВструктуру(Проект, "AcceptNewExternalData",	СвойстваПроекта);AcceptNewExternalData);
	ДобавитьСвойствоВСтруктуру(Проект, "AcceptNewExternalData", СвойстваПроекта);
	
	//    True if Project Server users can delegate tasks to other resources in the project. 
	// Read/write. Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AllowTaskDelegation", СвойстваПроекта);
	
	//    True if the actual, completed portion of a task that is scheduled before the status date
	// is moved to end at the status date.
	// Read/write. Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AndMoveCompleted", СвойстваПроекта);
	
	//    True if the remaining work on a task that is scheduled after the status date
	// is moved to start at the status date.
	// Read/write. Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AndMoveRemaining", СвойстваПроекта);
	
	// Gets the Application object.
	// Read-only. Application.
	ДобавитьСвойствоВструктуру(Проект, "Application", СвойстваПроекта);
	
	//    Gets or sets the way completed work is reported in team status messages.
	// Read/write PjTeamStatusCompletedWork.
	// pjBrokenDownByDay      (1) — Report work by day.
	// pjBrokenDownByWeek     (2) — Report work by week.
	// pjTotalForEntirePeriod (0) — Report total work for the entire period.
	//	
	ДобавитьСвойствоВструктуру(Проект, "AskForCompletedWork", СвойстваПроекта);
	
	//    True if new resources are automatically created as they are assigned.
	// False if Project prompts before creating new resources.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AutoAddResources", СвойстваПроекта);
	
	//    True if Project always calculates actual costs.
	// False if users can enter actual costs, and Project does not calculate actual costs.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AutoCalcCosts", СвойстваПроекта);
	
	//    Gets or sets whether the AutoFilter feature is turned on for a project.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AutoFilter", СвойстваПроекта);
	
	//    True if Project automatically links sequential tasks when you cut, move, or insert tasks.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AutoLinkTasks", СвойстваПроекта);
	
	//    True if Project automatically splits tasks into parts for work complete
	// and work remaining.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AutoSplitTasks", СвойстваПроекта);
	
	//    True if Project automatically updates the work and costs of resources
	// assigned to a task when the percent complete changes.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "AutoTrack", СвойстваПроекта);
	
	//    Gets a Calendars collection representing all base calendars in the active project.
	// Read-only Calendars.
	ДобавитьСвойствоВструктуру(Проект, "BaseCalendars", СвойстваПроекта);
	
	//    Gets date the specified baseline was last saved.
	// Read-only Variant. // Почему-то не берется. Исключение.
	ДобавитьСвойствоВструктуру(Проект, "BaselineSavedDate", СвойстваПроекта);
	
	//    Gets a DocumentProperties collection representing the built-in properties
	// of the document.
	// Read-only Object.
	ДобавитьСвойствоВструктуру(Проект, "BuiltinDocumentProperties",	СвойстваПроекта);
	
	//    Gets a Calendar object representing a calendar for the project.
	// Read-only Calendar.
	ДобавитьСвойствоВструктуру(Проект, "Calendar",				СвойстваПроекта);
	
	//    True if Project Professional can check in a project to Project Server.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "CanCheckIn", СвойстваПроекта);
	
	//    Gets the code name for the project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "CodeName", СвойстваПроекта);
	
	//    Gets a CommandBars collection that represents all the command bars in the project.
	// Read-only CommandBars.
	ДобавитьСвойствоВструктуру(Проект, "CommandBars", СвойстваПроекта);
	
	//    Gets the object that contains the embedded project.
	// Read-only Object.
	ДобавитьСвойствоВструктуру(Проект, "Container", СвойстваПроекта);
	
	//    Gets the date a project was created.
	// Read-only Variant.
	ДобавитьСвойствоВструктуру(Проект, "CreationDate", СвойстваПроекта);
	
	//    Project property for the three-character ISO standard currency code of the project.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "CurrencyCode", СвойстваПроекта);
	
	//    Sets or returns the number of digits following the decimal separator character
	// in currency values.
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "CurrencyDigits", СвойстваПроекта);
	
	//    Gets or sets the characters that denote currency values.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "CurrencySymbol", СвойстваПроекта);
	
	//    Gets or sets the location of the currency symbol. 
	// Read/write PjPlacement.
	//  pjBefore (0) — Before
	//  pjAfter (1) — After
	//  pjBeforeWithSpace (2) — Before with a space	
	//  pjAfterWithSpace  (3) — After with a space.
	ДобавитьСвойствоВструктуру(Проект, "CurrencySymbolPosition", СвойстваПроекта);
	
	//    Gets or sets the current date for a project.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "CurrentDate", СвойстваПроекта);
	
	//    Gets the name of the active filter for a project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "CurrentFilter", СвойстваПроекта);
	
	//    Gets the name of the active group for the active project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "CurrentGroup", СвойстваПроекта);
	
	// Gets the name of the active table for a project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "CurrentTable", СвойстваПроекта);
	
	//    Gets the name of the active view for a project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "CurrentView", СвойстваПроекта);
	
	//    Gets a DocumentProperties collection
	// representing the custom properties of the document.
	// Read-only Object.
	ДобавитьСвойствоВструктуру(Проект, "CustomDocumentProperties", СвойстваПроекта);
	
	//    Gets the project unique ID for a project stored in a database.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "DatabaseProjectUniqueID", СвойстваПроекта);
	
	//    Gets or sets the abbreviation for "day" that is displayed for values
	// such as durations, delays, slack, and work.
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "DayLabelDisplay", СвойстваПроекта);
	
	//    Gets or sets the number of days per month for tasks in a project
	// Read/write Double.
	ДобавитьСвойствоВструктуру(Проект, "DaysPerMonth", СвойстваПроекта);
	
	//    Gets or sets the default duration units.
	// Read/write PjUnit.
	//  pjMinute     (3) — Minute
	//  pjHour       (5) — Hour
	//  pjDay        (7) — Day
	//  pjWeek       (9) — Week
	//  pjMonthUnit (11) — Month.
	ДобавитьСвойствоВструктуру(Проект, "DefaultDurationUnits", СвойстваПроекта);
	
	//    Gets or sets the default method for calculating earned value for a project.
	// Read/write PjEarnedValueMethod.
	//  pjPercentComplete         (0) — Percent complete
	//  pjPhysicalPercentComplete (1) — Percent complete physically.
	ДобавитьСвойствоВструктуру(Проект, "DefaultEarnedValueMethod", СвойстваПроекта);
	
	//    True if new tasks are effort-driven by default.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "DefaultEffortDriven", СвойстваПроекта);
	
	//    Gets or sets the default finish time of the project.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "DefaultFinishTime", СвойстваПроекта);
	
	//    Gets or sets the default method used to accrue fixed task costs in the project.
	// Read/write PjAccrueAt.
	//  pjStart    (1) — Task accrues the resource cost when the task starts
	//  pjEnd      (2) — Task accrues the resource cost when the task ends
	//  pjProrated (3) — Task accrues the resource cost as the task progresses.
	ДобавитьСвойствоВструктуру(Проект, "DefaultFixedCostAccrual", СвойстваПроекта);
	
	//    Gets or sets the default overtime rate of pay for resources.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "DefaultResourceOvertimeRate", СвойстваПроекта);
	
	//    Gets or sets the default standard rate of pay for resources.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "DefaultResourceStandardRate", СвойстваПроекта);
	
	//    Gets or sets the default start time for the project.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "DefaultStartTime", СвойстваПроекта);
	
	//    Gets or sets the default task type.
	// Read/write PjTaskFixedType.
	//  pjFixedUnits    (0) — Fixed units
	//  pjFixedDuration (1) — Fixed duration
	//  pjFixedWork     (2) — Fixed work.
	ДобавитьСвойствоВструктуру(Проект, "DefaultTaskType", СвойстваПроекта);
	
	//    Gets or sets the default work units for the project. 
	// Read/write PjUnit.
	//  pjMinute     (3) — Minute
	//  pjHour       (5) — Hour
	//  pjDay        (7) — Day
	//  pjWeek       (9) — Week
	//  pjMonthUnit (11) — Month.
	ДобавитьСвойствоВструктуру(Проект, "DefaultWorkUnits", СвойстваПроекта);
	
	//    Gets a Tasks collection that contains a set of circular task dependencies,
	// if circular task references exist.
	// Read-only Tasks.
	ДобавитьСвойствоВструктуру(Проект, "DetectCycle", СвойстваПроекта);
	
	//    True if the summary task for a project is visible.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "DisplayProjectSummaryTask", СвойстваПроекта);
	
	//    Gets a DocumentLibraryVersions collection for the specified project.
	// Read-only DocumentLibraryVersions.
	ДобавитьСвойствоВструктуру(Проект, "DocumentLibraryVersions", СвойстваПроекта);
	
	//    Gets or sets the baseline for the earned values of tasks.
	// Read/write PjBaselines.
	//  pjBaseline    (0) — Baseline
	//  pjBaseline1   (1) — Baseline1
	//  pjBaseline2   (2) — Baseline2
	//  pjBaseline3   (3) — Baseline3
	//  pjBaseline4   (4) — Baseline4
	//  pjBaseline5   (5) — Baseline5
	//  pjBaseline6   (6) — Baseline6
	//  pjBaseline7   (7) — Baseline7
	//  pjBaseline8   (8) — Baseline8
	//  pjBaseline9   (9) — Baseline9
	//  pjBaseline10 (10) — Baseline10.
	ДобавитьСвойствоВструктуру(Проект, "EarnedValueBaseline", СвойстваПроекта);
	
	//    True if the actual work or actual overtime in a project is synchronized
	// with the actual work or actual overtime that has been submitted and updated
	// from the timesheet system.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "EnterpriseActualsSynched", СвойстваПроекта);
	
	//    True if timephased data is expanded to a readable format in the database.
	// False if timephased data is in a compressed binary format.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ExpandDatabaseTimephasedData", СвойстваПроекта);
	
	//    Gets or sets the color used to denote followed hyperlinks.
	// Read/write PjColor.
	//  pjBlack   (0) — Color is black
	//  pjRed     (1) — Color is red.
	//  pjYellow  (2) — Color is yellow
	//  pjLime    (3) — Color is lime
	//  pjAqua    (4) — Color is aqua
	//  pjBlue    (5) — Color is blue
	//  pjFuchsia (6) — Color is fuchsia
	//  pjWhite   (7) — Color is white
	//  pjMaroon  (8) — Color is maroon
	//  pjGreen   (9) — Color is green
	//  pjOlive  (10) — Color is olive
	//  pjNavy   (11) — Color is navy blue
	//  pjPurple (12) — Color is purple
	//  pjTeal   (13) — Color is teal
	//  pjGray   (14) — Color is gray
	//  pjSilver (15) — Color is silver
	//  pjColorAutomatic (16) — Color is selected automatically.
	ДобавитьСвойствоВструктуру(Проект, "FollowedHyperlinkColor", СвойстваПроекта);
	
	//    Gets or sets the color used to denote followed hyperlinks.
	// Read/write Long.
	ДобавитьСвойствоВструктуру(Проект, "FollowedHyperlinkColorEx", СвойстваПроекта);
	
	//    Gets the path and file name of a project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "FullName", СвойстваПроекта);
	
	//    True if a project has a password.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "HasPassword", СвойстваПроекта);
	
	//    True if tasks honor their constraint dates.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "HonorConstraints", СвойстваПроекта);
	
	//    Gets or sets the abbreviation for "hour" that is displayed for values
	// such as durations, delays, slack, and work. 
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "HourLabelDisplay", СвойстваПроекта);
	
	//    Gets or sets the number of hours per day for tasks in a project.
	// Read/write Double.
	ДобавитьСвойствоВструктуру(Проект, "HoursPerDay", СвойстваПроекта);
	
	//    Gets or sets the number of hours per week for tasks in a project.
	// Read/write Double.
	ДобавитьСвойствоВструктуру(Проект, "HoursPerWeek", СвойстваПроекта);
	
	//    Gets or sets the color used to denote unfollowed hyperlinks.
	// Read/write PjColor.
	//  pjBlack   (0) — Color is black
	//  pjRed     (1) — Color is red.
	//  pjYellow  (2) — Color is yellow
	//  pjLime    (3) — Color is lime
	//  pjAqua    (4) — Color is aqua
	//  pjBlue    (5) — Color is blue
	//  pjFuchsia (6) — Color is fuchsia
	//  pjWhite   (7) — Color is white
	//  pjMaroon  (8) — Color is maroon
	//  pjGreen   (9) — Color is green
	//  pjOlive  (10) — Color is olive
	//  pjNavy   (11) — Color is navy blue
	//  pjPurple (12) — Color is purple
	//  pjTeal   (13) — Color is teal
	//  pjGray   (14) — Color is gray
	//  pjSilver (15) — Color is silver
	//  pjColorAutomatic (16) — Color is selected automatically.
	ДобавитьСвойствоВструктуру(Проект, "HyperlinkColor", СвойстваПроекта);
	
	//    Gets or sets a hexadecimal representation of the color used to denote
	// unfollowed hyperlinks.
	// Read/write Long.
	ДобавитьСвойствоВструктуру(Проект, "HyperlinkColorEx", СвойстваПроекта);
	
	//    Gets the identification number of a project.
	// Read-only Long.
	ДобавитьСвойствоВструктуру(Проект, "ID", СвойстваПроекта);
	
	//    Gets the index of a Project object in the containing Projects collection.
	// Read-only Variant.
	ДобавитьСвойствоВструктуру(Проект, "Index", СвойстваПроекта);
	
	//    Gets whether the checkout message bar is visible.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "IsCheckoutMsgBarVisible", СвойстваПроекта);
	
	//    Gets whether the Check Out button is visible in the Backstage view.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "IsCheckoutOSVisible", СвойстваПроекта);
	
	//    Gets or sets a value that indicates whether the project is a template.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "IsTemplate", СвойстваПроекта);
	
	//    True if task scheduling respects the current calendar when a task is converted
	// from manual to automatic; otherwise, False. 
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "KeepTaskOnNearestWorkingTimeWhenMadeAutoScheduled", СвойстваПроекта);
	
	//    Gets the date a project was last printed.
	// Read-only Variant.
	ДобавитьСвойствоВструктуру(Проект, "LastPrintedDate", СвойстваПроекта);
	
	//    Gets the date a project was last saved.
	// Read-only Variant.
	ДобавитьСвойствоВструктуру(Проект, "LastSaveDate", СвойстваПроекта);
	
	//    Gets the name of the user who last saved a project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "LastSavedBy", СвойстваПроекта);
	
	//    True if all resources in the project are leveled.
	// False if only overallocated resources within specified dates are leveled.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "LevelEntireProject", СвойстваПроекта);
	
	//    Gets or sets the starting date of a range in which overallocated resources
	// are leveled. The default is the project start date or the last entered date value.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "LevelFromDate", СвойстваПроекта);
	
	//    Gets or sets the ending date of a range in which overallocated resources are leveled.
	// The default is the project finish date or the last entered date value.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "LevelToDate", СвойстваПроекта);
	
	//    True if predecessor and successor task links are maintained
	// when a task is converted from automatic to manual; otherwise, False.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ManuallyScheduledTasksAutoRespectLinks", СвойстваПроекта);
	
	//    Gets a List object representing the list of data maps in the project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "MapList", СвойстваПроекта);
	
	//    Gets or sets the abbreviation for "minute" that is displayed for values
	// such as durations, delays, slack, and work.
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "MinuteLabelDisplay", СвойстваПроекта);
	
	//    Gets or sets the abbreviation for "month" that is displayed for values
	// such as durations, delays, slack, and work.
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "MonthLabelDisplay", СвойстваПроекта);
	
	//    True if a task that is scheduled after the status date has actual progress
	// entered against it and the actual, completed portion of the task is moved
	// so the completed work ends on the status date.
	// Read/Write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "MoveCompleted", СвойстваПроекта);
	
	//    True if the remaining portion of a task that is scheduled before
	// the status date is moved to start at the status date.
	// Read/Write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "MoveRemaining", СвойстваПроекта);
	
	//    True if Project calculates multiple critical paths for the project.
	// False if only one critical path is calculated.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "MultipleCriticalPaths", СвойстваПроекта);
	
	//    Gets the name of a Project object.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "Name", СвойстваПроекта);
	
	//    True if new tasks are created as manually scheduled tasks.
	// False if new tasks are automatically scheduled.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "NewTasksCreatedAsManual", СвойстваПроекта);
	
	//    True if new tasks in the active project have estimated durations.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "NewTasksEstimated", СвойстваПроекта);
	
	//    Gets the number of resources in a project, not including blank entries.
	// Read-only Long.
	ДобавитьСвойствоВструктуру(Проект, "NumberOfResources", СвойстваПроекта);
	
	//    Gets the number of tasks in a project, not including blank entries.
	// Read-only Long.
	ДобавитьСвойствоВструктуру(Проект, "NumberOfTasks", СвойстваПроекта);
	
	//    Gets a Tasks collection representing the children of a task
	// in the outline structure.
	// Read-only Tasks.
	ДобавитьСвойствоВструктуру(Проект, "OutlineChildren", СвойстваПроекта);
	
	//    Gets an OutlineCodes collection of all outline codes defined for resources and tasks
	// in the project.
	// Read-only OutlineCodes.
	ДобавитьСвойствоВструктуру(Проект, "OutlineCodes", СвойстваПроекта);
	
	//    Gets the parent of the Project object.
	// Read-only Object.
	ДобавитьСвойствоВструктуру(Проект, "Parent", СвойстваПроекта);
	
	//    Gets the path of the open project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "Path", СвойстваПроекта);
	
	//    Gets or sets the type of characters used to display phonetic information.
	// Read/write PjPhoneticType.
	//  pjKatakanaHalf (0) — Half-width Katakana characters
	//  pjKatakana     (1) — Katakana characters
	//  pjHiragana     (2) — Hiragana characters.
	ДобавитьСвойствоВструктуру(Проект, "PhoneticType", СвойстваПроекта);
	
	//    Gets or sets the finish date for a project.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "ProjectFinish", СвойстваПроекта);
	
	//    Gets or sets the name of the XML schema being used by the Project Guide.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "ProjectGuideContent", СвойстваПроекта);
	
	//    Gets or sets the Project Guide functional layout page for the specified project.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "ProjectGuideFunctionalLayoutPage", СвойстваПроекта);
	
	//    Gets or sets an XML string representing the save buffer of the Project Guide.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "ProjectGuideSaveBuffer", СвойстваПроекта);
	
	//    True if the Project Guide uses the default content.
	// False if you want to use custom content for the Project Guide.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ProjectGuideUseDefaultContent", СвойстваПроекта);
	
	//    True if Project uses the default Project Guide.
	// False if you are customizing the Project Guide.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ProjectGuideUseDefaultFunctionalLayoutPage", СвойстваПроекта);
	
	//    Gets the prefix of the project name of the specified project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "ProjectNamePrefix", СвойстваПроекта);
	
	//    Gets or sets the notes for the project.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "ProjectNotes", СвойстваПроекта);
	
	//    True if Project Server is used for tracking the specified project.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ProjectServerUsedForTracking", СвойстваПроекта);
	
	//    Gets or sets the start date for a project.
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "ProjectStart", СвойстваПроекта);
	
	//    Gets a Task object representing the project summary task for the active project.
	// Read-only Task.
	ДобавитьСвойствоВструктуру(Проект, "ProjectSummaryTask", СвойстваПроекта);
	
	//    True if a project has read-only access.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ReadOnly", СвойстваПроекта);
	
	//    True if the project should be opened with read-only access.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ReadOnlyRecommended", СвойстваПроекта);
	
	//    True if Project removes user information from revisions and the project Properties
	// dialog box upon saving a document.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "RemoveFileProperties", СвойстваПроекта);
	
	//    Gets the collection of custom reports in the project.
	// Read-only Reports.
	ДобавитьСвойствоВструктуру(Проект, "Reports", СвойстваПроекта);
	
	//    Gets a List object representing all resource filters in the project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "ResourceFilterList", СвойстваПроекта);
	
	//    Gets a Filters collection that contains the resource filters of the project.
	// Read-only Filters.
	ДобавитьСвойствоВструктуру(Проект, "ResourceFilters", СвойстваПроекта);
	
	//    Gets a List object representing the resource groups in the active project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "ResourceGroupList", СвойстваПроекта);
	
	//    Gets a ResourceGroups collection that contains all the resource-based group definitions
	// in the project.
	// Read-only ResourceGroups.
	ДобавитьСвойствоВструктуру(Проект, "ResourceGroups", СвойстваПроекта);
	
	// Gets a ResourceGroups2 collection that represents all of the resource groups
	// based on Group2 objects.
	// Read-only ResourceGroups2.
	ДобавитьСвойствоВструктуру(Проект, "ResourceGroups2", СвойстваПроекта);
	
	//    Gets the name of the enterprise resource pool that a project uses
	// in Project Professional.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "ResourcePoolName", СвойстваПроекта);
	
	//    Gets a Resources collection representing the resources in a Project.
	// Read-only Object.
	ДобавитьСвойствоВструктуру(Проект, "Resources", СвойстваПроекта);
	
	//    Gets a List object representing all resource tables in the project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "ResourceTableList", СвойстваПроекта);
	
	//    Gets a Tables collection that contains the resource tables of the project.
	// Read-only Tables.
	ДобавитьСвойствоВструктуру(Проект, "ResourceTables", СвойстваПроекта);
	
	//    Gets a List object representing all resource views in the active project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "ResourceViewList", СвойстваПроекта);
	
	//    Gets the number of times a project has been saved.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "RevisionNumber", СвойстваПроекта);
	
	//    True if a project has not changed since it was last saved.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "Saved", СвойстваПроекта);
	
	//    True if Project calculates the project schedule forward from the start date.
	// False if the schedule is calculated backward from the finish date.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ScheduleFromStart", СвойстваПроекта);
	
	//    Gets or sets the way Project Professional users are identified to Project Server.
	// Read/write PjAuthentication.
	//  pjUserName           (0) — Authenticated by user name
	//  pjWindowsUserAccount (1) — Authenticated by user Windows account.
	ДобавитьСвойствоВструктуру(Проект, "ServerIdentification", СвойстваПроекта);
	
	//    Gets the URL of the Project Web App instance with which Project Professional
	// is connected. For a synchronized SharePoint task list, gets or sets an arbitrary value
	// that has no effect on the project.
	// Read/write String.
	ДобавитьСвойствоВструктуру(Проект, "ServerURL", СвойстваПроекта);
	
	//    Gets a SharedWorkspace object that represents the document workspace for the project.
	// Read-only SharedWorkspace.
	ДобавитьСвойствоВструктуру(Проект, "SharedWorkspace", СвойстваПроекта);
	
	//    Gets or sets how much slack causes a task to be displayed as a critical task.
	// Read/write Long.
	ДобавитьСвойствоВструктуру(Проект, "ShowCriticalSlack", СвойстваПроекта);
	
	//    True if the Links between Projects dialog box appears when a project containing
	// cross-project links is opened.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ShowCrossProjectLinksInfo", СвойстваПроекта);
	
	//    True if task durations in the project are displayed with the estimated character.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ShowEstimatedDuration", СвойстваПроекта);
	
	//    True if predecessor tasks linked from an external project should be displayed.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ShowExternalPredecessors", СвойстваПроекта);
	
	//    True if successor tasks linked from an external project should be displayed.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ShowExternalSuccessors", СвойстваПроекта);
	
	//    True if task suggestions in the active project are displayed; otherwise, False.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ShowTaskSuggestions", СвойстваПроекта);
	
	//    True if task warnings in the active project are displayed; otherwise, False.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "ShowTaskWarnings", СвойстваПроекта);
	
	//    True if a time value should be separated from its time label by a space.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "SpaceBeforeTimeLabels", СвойстваПроекта);
	
	//    True if edits to total actual cost are spread to the status date,
	// or to the current date if the status date is "NA".
	// False if edits are spread to the calculated stop date of the task.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "SpreadCostsToStatusDate", СвойстваПроекта);
	
	//    True if edits to total task percent complete are spread to the status date,
	// or to the current date if the status date is "NA".
	// False if edits are spread to the calculated stop date of the task.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "SpreadPercentCompleteToStatusDate", СвойстваПроекта);
	
	//    True if new tasks start on the current date.
	// False if new tasks start on the project start date.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "StartOnCurrentDate", СвойстваПроекта);
	
	//    Gets or sets the first day of the week for the project.
	// Read/write PjWeekday.
	//  pjSunday    (1) — Sunday
	//  pjMonday    (2) — Monday
	//  pjTuesday   (3) — Tuesday
	//  pjWednesday (4) — Wednesday
	//  pjThursday  (5) — Thursday
	//  pjFriday    (6) — Friday
	//  pjSaturday  (7) — Saturday.
	ДобавитьСвойствоВструктуру(Проект, "StartWeekOn", СвойстваПроекта);
	
	//    Gets or sets the month number for the start of the fiscal year for the project.
	// Read/write PjMonth.
	//  pjJanuary (1) — January
	//  pjFebruary (2) — February
	//  pjMarch (3) — March
	//  pjApril (4) — April
	//  pjMay (5) — May
	//  pjJune (6) — June
	//  pjJuly (7) — July
	//  pjAugust (8) — August
	//  pjSeptember (9) — September
	//  pjOctober (10) — October
	//  pjNovember (11) — November
	//  pjDecember (12) — December.
	ДобавитьСвойствоВструктуру(Проект, "StartYearIn", СвойстваПроекта);
	
	//    Gets or sets the current status date for the project.
	// If there is no status date, returns "NA".
	// Read/write Variant.
	ДобавитьСвойствоВструктуру(Проект, "StatusDate", СвойстваПроекта);
	
	//    Gets a Subprojects collection representing subprojects in the master project.
	// Read-only Subprojects.
	ДобавитьСвойствоВструктуру(Проект, "Subprojects", СвойстваПроекта);
	
	//    Gets or sets the number of task errors associated with a project.
	// Read/write Long.
	ДобавитьСвойствоВструктуру(Проект, "TaskErrorCount", СвойстваПроекта);
	
	//    Gets a List object representing all task filters in the project.
	// Read-only List .
	ДобавитьСвойствоВструктуру(Проект, "TaskFilterList", СвойстваПроекта);
	
	//    Gets a Filters collection of the task filters in the project.
	// Read-only Filters.
	ДобавитьСвойствоВструктуру(Проект, "TaskFilters", СвойстваПроекта);
	
	//    Gets a List object representing the task groups in the active project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "TaskGroupList", СвойстваПроекта);
	
	//    Gets a TaskGroups collection representing all the task-based Group definitions
	// in the project.
	// Read-only TaskGroups.
	ДобавитьСвойствоВструктуру(Проект, "TaskGroups", СвойстваПроекта);
	
	//    Gets a TaskGroups2 collection that represents all the task-based Group2 definitions
	// in the specified project.
	// Read-only TaskGroups2.
	ДобавитьСвойствоВструктуру(Проект, "TaskGroups2", СвойстваПроекта);
	
	//    Gets a Tasks collection representing the tasks in the project.
	// Read-only Tasks.
	ДобавитьСвойствоВструктуру(Проект, "Tasks", СвойстваПроекта);
	
	//    Gets a List object representing all task tables in the project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "TaskTableList", СвойстваПроекта);
	
	//    Gets a Tables collection representing the task tables in the project.
	// Read-only Tables.
	ДобавитьСвойствоВструктуру(Проект, "TaskTables", СвойстваПроекта);
	
	//    Gets a List object representing all task views in the project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "TaskViewList", СвойстваПроекта);
	
	//    Gets the name of the template associated with a project.
	// Read-only String.
	ДобавитьСвойствоВструктуру(Проект, "Template", СвойстваПроекта);
	
	//    Gets or sets the tracking method used by Project Server for the project.
	// Read/write PjProjectServerTrackingMethod.
	//  pjTrackingMethodDefault           (0) — Use the default tracking method
	//  pjTrackingMethodSpecifyHours      (1) — Tracking is by reported hours worked per period
	//  pjTrackingMethodPercentComplete   (2) — Tracking is by percent complete
	//  pjTrackingMethodTotalAndRemaining (3) — Tracking is by total work completed and estimated remaining work.
	ДобавитьСвойствоВструктуру(Проект, "TrackingMethod", СвойстваПроекта);
	
	//    Gets the type of a project.
	// Read-only PjProjectType.
	//  pjProjectTypeUnsaved                (0) — Project type is unsaved.
	//  pjProjectTypeNonEnterprise           (1) — Project type is nonenterprise.
	//  pjProjectTypeEnterpriseCheckedOut     (2) — Project type is enterprise checked out.
	//  pjProjectTypeEnterpriseReadOnly        (3) — Project type is enterprise read-only.
	//  pjProjectTypeEnterpriseGlobalCheckedOut (4) — Project type is enterprise global checked out.
	//  pjProjectTypeEnterpriseGlobalInMemory    (5) — Project type is enterprise global in memory.
	//  pjProjectTypeEnterpriseGlobalLocal        (6) — Project type is enterprise global local.
	//  pjProjectTypeEnterpriseResourcesCheckedOut (7) — Project type is enterprise resources checked out.
	ДобавитьСвойствоВструктуру(Проект, "Type", СвойстваПроекта);
	
	//    True if hyperlinks are underlined.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "UnderlineHyperlinks", СвойстваПроекта);
	
	//    Gets the unique identification number of the project,
	// which is actually the UniqueID value of the project summary task.
	// Read-only Long.
	ДобавитьСвойствоВструктуру(Проект, "UniqueID", СвойстваПроекта);
	
	//    True if a fiscal year is determined by the year of the first month
	// of that fiscal year. False if determined by the last month of the fiscal year.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "UseFYStartYear", СвойстваПроекта);
	
	//    True if the user directly opens or creates the project.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "UserControl", СвойстваПроекта);
	
	//    True if the Microsoft Visual Basic for Applications project is digitally signed.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "VBASigned", СвойстваПроекта);
	
	//    Gets a VBProject object that represents the Microsoft Visual Basic project.
	// Read-only VBProject.
	ДобавитьСвойствоВструктуру(Проект, "VBProject", СвойстваПроекта);
	
	//    Gets the List object for the project.
	// Read-only List.
	ДобавитьСвойствоВструктуру(Проект, "ViewList", СвойстваПроекта);
	
	//    Gets a Views collection representing the views of the project.
	// Read-only Views.
	ДобавитьСвойствоВструктуру(Проект, "Views", СвойстваПроекта);
	
	//    Gets a ViewsCombination collection representing the combination views of the project.
	// Read-only ViewsCombination.
	ДобавитьСвойствоВструктуру(Проект, "ViewsCombination", СвойстваПроекта);
	
	//    Gets a ViewsSingle collection representing the single views of the project.
	// Read-only ViewsSingle.
	ДобавитьСвойствоВструктуру(Проект, "ViewsSingle", СвойстваПроекта);
	
	//    True if a work breakdown structure (WBS) code is automatically generated
	// for new tasks in the project.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "WBSCodeGenerate", СвойстваПроекта);
	
	//    True if an edited work breakdown structure (WBS) code is verified to be unique.
	// Read/write Boolean.
	ДобавитьСвойствоВструктуру(Проект, "WBSVerifyUniqueness", СвойстваПроекта);
	
	//    Gets or sets the abbreviation for "week" that is displayed for values
	// such as durations, delays, slack, and work.
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "WeekLabelDisplay", СвойстваПроекта);
	
	//    Gets a Windows collection representing the open windows in the project.
	// Read-only Windows.
	ДобавитьСвойствоВструктуру(Проект, "Windows", СвойстваПроекта);
	
	//    Gets a Windows2 collection representing the open windows in the project.
	// Read-only Windows2.
	ДобавитьСвойствоВструктуру(Проект, "Windows2", СвойстваПроекта);
	
	//    True if a password is required to open a project for read/write access.
	// Read-only Boolean.
	ДобавитьСвойствоВструктуру(Проект, "WriteReserved", СвойстваПроекта);
	
	//    Gets or sets how the year label displays in rates.
	// Read/write Integer.
	ДобавитьСвойствоВструктуру(Проект, "YearLabelDisplay", СвойстваПроекта);
	
	Возврат СвойстваПроекта;
КонецФункции

#КонецОбласти

Функция ПолучитьОбъектXDTO(ПространствоИмен, Тип) Экспорт
	
	Попытка
		
		ОбъектДанныеТип = ФабрикаXDTO.Тип(ПространствоИмен, Тип );
		ОбъектДанные = ФабрикаXDTO.Создать(ОбъектДанныеТип);    
		
	Исключение
		
		Описание = ПодробноеПредставлениеОшибки(ИнформацияОбОшибке());
		ТекстОшибки = "Не удалось получить Тип XDTO «" + 
		Тип + "» в пакете «"+ПространствоИмен+"» по причине: " + Описание;
		
	КонецПопытки;
	
	
	Возврат ОбъектДанные;
	
КонецФункции

Процедура СообщитьОбУдаленныхРемонтах(МассивУдаленныхРемонтов)
	
	КоличествоУдаленных = МассивУдаленныхРемонтов.Количество();
	
	Если КоличествоУдаленных > 0 Тогда
		ШаблонСообщения = НСтр("ru = 'При редактировании в MS Project из файла '");
		Если КоличествоУдаленных = 1 Тогда
			ШаблонСообщения = ШаблонСообщения + НСтр("ru = 'был удален 1 ремонт.'");
		Иначе
			Если КоличествоУдаленных<5 Тогда
				ШаблонСообщения = ШаблонСообщения + НСтр("ru = 'было удалено %1 ремонта.'");
			Иначе
				ШаблонСообщения = ШаблонСообщения + НСтр("ru = 'было удалено %1 ремонтов.'");
			КонецЕсли;
			ТекстСообщения = СтрШаблон(ШаблонСообщения,КоличествоУдаленных);
		КонецЕсли;
		ОбщегоНазначенияКлиент.СообщитьПользователю(ТекстСообщения);
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти

#КонецЕсли