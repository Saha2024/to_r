
#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область СлужебныеПроцедурыИФункции

// Настройки размещения в панели отчетов. (См. ВариантыОтчетовПереопределяемый.НастроитьВариантыОтчетов).
//
// Параметры:
//   Настройки - Коллекция - Передается "как есть" из ВариантыОтчетовПереопределяемый.НастроитьВариантыОтчетов().
//       Может использоваться для получения настроек варианта этого отчета при помощи функции ВариантыОтчетов.ОписаниеВарианта().
//   НастройкиОтчета - СтрокаДереваЗначений - Настройки этого отчета,
//       уже сформированные при помощи функции ВариантыОтчетов.ОписаниеОтчета() и готовые к изменению.
//       См. "Свойства для изменения" процедуры ВариантыОтчетовПереопределяемый.НастроитьВариантыОтчетов().
//
Процедура НастроитьВариантыОтчета(Настройки, НастройкиОтчета) Экспорт
	
	НастройкиОтчета.ОпределитьНастройкиФормы = Истина;
	
	НастройкиОтчета.ФункциональныеОпции.Добавить("торо_ИспользоватьСметыРемонта");
	ВариантыОтчетов.УстановитьРежимВыводаВПанеляхОтчетов(Настройки, НастройкиОтчета, Истина);

	ОписаниеВарианта = ВариантыОтчетов.ОписаниеВарианта(Настройки, НастройкиОтчета, "Основной");
	ОписаниеВарианта.ВидимостьПоУмолчанию = Истина;
	ОписаниеВарианта.Наименование = НСтр("ru = 'Статусы ремонтов по заявкам'");
	ОписаниеВарианта.Описание = НСтр("ru = 'История изменения статусов ремонтов по заявкам на ремонты.'");
	
	// Выдает ошибку на версии платформы 8.3.22
	ОписаниеВарианта = ВариантыОтчетов.ОписаниеВарианта(Настройки, НастройкиОтчета, "ТаблицаСИерархией");
	ОписаниеВарианта.ВидимостьПоУмолчанию = Ложь; 
	ОписаниеВарианта.Включен = Ложь;
	ОписаниеВарианта.Наименование = НСтр("ru = 'Статусы ремонтов по заявкам (с иерархией)'");
	ОписаниеВарианта.Описание = НСтр("ru = 'История изменения статусов ремонтов по заявкам на ремонты (с иерархией объектов ремонта).'");
	
	ОписаниеВарианта = ВариантыОтчетов.ОписаниеВарианта(Настройки, НастройкиОтчета, "ПоБригадам");
	ОписаниеВарианта.ВидимостьПоУмолчанию = Истина;
	ОписаниеВарианта.Наименование = НСтр("ru = 'Статусы ремонтов по заявкам (по бригадам)'");
	ОписаниеВарианта.Описание = НСтр("ru = 'История изменения статусов ремонтов по заявкам на ремонты в разрезе исполнителей.'");
	
	// Выдает ошибку на версии платформы 8.3.22
	ОписаниеВарианта = ВариантыОтчетов.ОписаниеВарианта(Настройки, НастройкиОтчета, "ТаблицаПоБригадамСИерархией");
	ОписаниеВарианта.ВидимостьПоУмолчанию = Ложь;
	ОписаниеВарианта.Включен = Ложь;
	ОписаниеВарианта.Наименование = НСтр("ru = 'Статусы ремонтов по заявкам (по бригадам, с иерархией)'");
	ОписаниеВарианта.Описание = НСтр("ru = 'История изменения статусов ремонтов по заявкам на ремонты в разрезе исполнителей (с иерархией объектов ремонта).'");

КонецПроцедуры

#КонецОбласти

#КонецЕсли