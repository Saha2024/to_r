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
	
	НастройкиОтчета.Размещение.Вставить(ВариантыОтчетовКлиентСервер.ИдентификаторНачальнойСтраницы());
	ВариантыОтчетов.УстановитьРежимВыводаВПанеляхОтчетов(Настройки, НастройкиОтчета, Истина);

	ОписаниеВарианта = ВариантыОтчетов.ОписаниеВарианта(Настройки, НастройкиОтчета, "Номенклатура");
	ОписаниеВарианта.ВидимостьПоУмолчанию = Истина;
	ОписаниеВарианта.Наименование = НСтр("ru = 'План-фактный анализ затрат МТО'");
	ОписаниеВарианта.Описание = НСтр("ru = 'План и факт материальных затрат за период в количественном и суммовом выражении.'");
	
	ОписаниеВарианта = ВариантыОтчетов.ОписаниеВарианта(Настройки, НастройкиОтчета, "НоменклатураИерархия");
	ОписаниеВарианта.ВидимостьПоУмолчанию = Истина;
	ОписаниеВарианта.Наименование = НСтр("ru = 'План-фактный анализ затрат МТО (с иерархией)'");
	ОписаниеВарианта.Описание = НСтр("ru = 'План и факт материальных затрат за период в количественном и суммовом выражении (с иерархией объектов ремонта).'");

КонецПроцедуры

#КонецОбласти

#КонецЕсли