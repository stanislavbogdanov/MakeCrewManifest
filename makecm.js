/************************************************************************/
/* MAKECM.JS Создание файлов Crew Manifest для проекта CrewTablet       */
/* Автор: Станислав О. Богданов (c) 2014                                */
/************************************************************************/

/* ИСТОРИЯ ВЕРСИЙ *******************************************************/
/* 24.10.2014 Ver. 1.01                                                 */
/* Исправлена критическая ошибка при разборе SSIM и создании таблицы    */
/* TIMETABLE: нумерация месяцев в году в FoxPro 1-12, WScript 0-11.     */
/************************************************************************/
/* 21.08.2014 Ver. 1.00                                                 */
/* Создание CrewManifestFull для рейсов команды тестеров на дату        */
/* по данным из БД TRUCK планирования экипажей, расписания рейсов       */
/* в формате SSIM.                                                      */
/************************************************************************/

// Путь к файлам DBF-таблиц и расписания SSIM
var sDBSourcePath = "Tables";
// Путь к файлам Crew Manifest
var sManifestPath = "CrewManifest";
// Имя файла SSIM
var sSSIMFileName = "Current_SSIM";
// ConnectionString для объектов ADODB
var sConnect = "provider=vfpoledb.1;;data source=" + sDBSourcePath;

// Имена файлов манифестов для прямого и обратного рейсов
var sManifestFName, sManifestFName2;

// Поля Crew Manifest
var fieldStaffId;
var fieldName;
var fieldPositionA;
var fieldLanguage;
var fieldFlight;           
var fieldPositionN;
var fieldDeptDateTimeUTC;   
var fieldArvlDateTimeUTC;   
var fieldDeptArp;           
var fieldArvlArp;           
var fieldEmail;
var fieldContact;
var fieldNationality;
var fieldAircraftType;       
var fieldAircraftRegNum;

// Индексы в массиве, содержащем сведения о членах бригады
var indexStaffId       = 0;
var indexName          = 1;
var indexPositionA     = 2;
var indexLanguage      = 3;
var indexPositionN     = 4;
var indexEmail         = 5;
var indexContact       = 6;
var indexNationality   = 7;
var indexAircraftType  = 8;
var indexChiefPriority = 9;

function ShiftFreq(src, shift)
{
  var x, i;
  var dstarr = ["1","2","3","4","5","6","7"];
  var dststr = "";
 
  for (x=0; x<7; x++) {
    i = (x+shift) % 7;
    if (src.charAt(x) == ' ') dstarr[i] = ' ';
  }
  for (x=0; x<7; x++) dststr += dstarr[x];
  return(dststr);
}

function ParseSSIMDate(s)
{
  var d, m, y;
  var SSIMDate;
 
  d = s.substring(0,2);
  y = "20" + s.substring(5);
  switch (s.substring(2,5)) {
    case "JAN": m = 0; break;
    case "FEB": m = 1; break;
    case "MAR": m = 2; break;
    case "APR": m = 3; break;
    case "MAY": m = 4; break;
    case "JUN": m = 5; break;
    case "JUL": m = 6; break;
    case "AUG": m = 7; break;
    case "SEP": m = 8; break;
    case "OCT": m = 9; break;
    case "NOV": m = 10; break;
    case "DEC": m = 11; break;
    default: m = 0; break;
  }
  SSIMDate = new Date(y, m, d);
  return (SSIMDate);
}

function ReadSSIM(sSSIMPath)
{
  var FSO;
  var ForReading = 1;
  var SSIM;
  var TristateFalse = 0; // ASCII mode
  var s = "";
  var FlightNumber, ServiceType, EffDate, DisDate, Freq;
  var DeptArp, DeptTimeUTC, DeptTimeUTCInMin, DeptTimeZone, DeptTimeZoneInMin;
  var ArvlArp, ArvlTimeUTC, ArvlTimeUTCInMin, ArvlTimeZone, ArvlTimeZoneInMin;
  var AircraftType;
  var EffDateLoc, DisDateLoc;
  var FlightDurationInMin, A_EffDateLoc, A_DisDateLoc;
  var oConn, sReq;
  var ael;
 
  FSO = WScript.CreateObject("Scripting.FileSystemObject");
  SSIM = FSO.OpenTextFile(sDBSourcePath+"\\"+sSSIMFileName, ForReading, false, TristateFalse);

  sReq = "CREATE TABLE TIMETABLE (FLIGHTNUM Integer, SRVCTYPE Char(1), ";
  sReq += "EFFDATE_U DateTime, DISDATE_U DateTime, FREQ_U Char(7), ";
  sReq += "DTIME_U Integer, DTIMEZONE Integer, ";
  sReq += "ATIME_U Integer, ATIMEZONE Integer, ";
  sReq += "EFFDATE_L DateTime, DISDATE_L DateTime, FREQ_L Char(7), DTIME_L Integer, ";
  sReq += "DARP Char(3), AARP Char(3), ";
  sReq += "AEFFDATE_L DateTime, ADISDATE_L DateTime, AFREQ_L Char(7), ATIME_L Integer, ";
  sReq += "AIRCTYPE Char (3)";
  sReq += ")";
   
  oConn = WScript.CreateObject("ADODB.Connection");
  oConn.Open(sConnect);
  oConn.Execute(sReq);
 
  while (!SSIM.AtEndOfStream) {
    s = SSIM.ReadLine();
    if (s.charAt(0) == "3") {  // Flight Leg Record (Record Type 3)
      FlightNumber = s.substring(5,9);
      ServiceType = s.substring(13,14);
      EffDate = ParseSSIMDate(s.substring(14,21));
      DisDate = ParseSSIMDate(s.substring(21,28));
      Freq = s.substring(28,35);
     
      DeptArp = s.substring(36,39);
      DeptTimeUTC = s.substring(39,43);
      DeptTimeUTCInMin = Math.floor(StrToInt(DeptTimeUTC)/100)*60 + StrToInt(DeptTimeUTC%100);
      DeptTimeZone = s.substring(47,52);
      DeptTimeZoneInMin = StrToInt(DeptTimeZone.charAt(0) + (StrToInt(DeptTimeZone.substring(1,3))*60+StrToInt(DeptTimeZone.substring(3))));
     
      ArvlArp = s.substring(54,57);
      ArvlTimeUTC = s.substring(61,65);
      ArvlTimeUTCInMin = StrToInt(ArvlTimeUTC.substring(0,2))*60+StrToInt(ArvlTimeUTC.substring(2));
      ArvlTimeZone = s.substring(65,70);
      ArvlTimeZoneInMin = StrToInt(ArvlTimeZone.charAt(0) + (StrToInt(ArvlTimeZone.substring(1,3))*60+StrToInt(ArvlTimeZone.substring(3))));
      AircraftType = s.substring(72,75);
     
      sReq = "INSERT INTO TIMETABLE (";
      sReq += "FLIGHTNUM, SRVCTYPE, EFFDATE_U, DISDATE_U, FREQ_U, DTIME_U, DTIMEZONE, ATIME_U, ATIMEZONE, ";
      sReq += "EFFDATE_L, DISDATE_L, FREQ_L, DTIME_L, DARP, AARP, AEFFDATE_L, ADISDATE_L, AFREQ_L, ATIME_L, AIRCTYPE";
      sReq += ") ";
      sReq += "VALUES (";
      // FLIGHTNUM    Integer        Номер рейса
      sReq += FlightNumber + ", ";
      // SRVCTYPE    Char(1)        "J" - регулярный рейс
      sReq += "'" + ServiceType + "'" + ", ";
      // EFFDATE_U    DateTime    Начало навигации UTC (дата вылета)
      sReq += "DATETIME(" + EffDate.getFullYear() + "," + (EffDate.getMonth()+1) + "," + EffDate.getDate() + ")";
      sReq += ", ";
      // DISDATE_U    DateTime    Конец навигации UTC (дата вылета)
      sReq += "DATETIME(" + DisDate.getFullYear() + "," + (DisDate.getMonth()+1) + "," + DisDate.getDate() + ")" ;
      sReq += ", ";
      // FREQ_U        Char(7)        Дни недели вылета UTC
      sReq += "'" + Freq + "'" + ", ";
      // DTIME_U    Integer        Время вылета UTC (в минутах), всегда положительное
      sReq += DeptTimeUTCInMin + ", ";
      // DTIMEZONE    Integer        Часовой пояс аэропорта вылета (в минутах)
      sReq += DeptTimeZoneInMin + ", ";
      // ATIME_U    Integer        Время прилета UTC (в минутах), всегда положительное
      sReq += ArvlTimeUTCInMin + ", ";
      // ATIMEZONE    Integer        Часовой пояс аэропорта прилета (в минутах)
      sReq += ArvlTimeZoneInMin + ", ";

      // EFFDATE_L    DateTime    Начало навигации LOC (дата вылета)
      EffDateLoc = new Date(EffDate);
      EffDateLoc.setMinutes(DeptTimeUTCInMin + DeptTimeZoneInMin);
      sReq += "DATETIME(" + EffDateLoc.getFullYear() + "," + (EffDateLoc.getMonth()+1) + "," + EffDateLoc.getDate() + ")";
      sReq += ", ";
      // DISDATE_L    DateTime    Конец навигации LOC (дата вылета)
      DisDateLoc = new Date(DisDate);
      DisDateLoc.setMinutes(DeptTimeUTCInMin + DeptTimeZoneInMin);
      sReq += "DATETIME(" + DisDateLoc.getFullYear() + "," + (DisDateLoc.getMonth()+1) + "," + DisDateLoc.getDate() + ")";
      sReq += ", ";
      // FREQ_L        Char(7)        Дни недели вылета LOC
      sReq += "'" + ShiftFreq(Freq, Math.floor((DeptTimeUTCInMin+DeptTimeZoneInMin)/(24*60))) + "'";
      sReq += ", ";
      // DTIME_L    Integer        Время вылета LOC (в минутах)
      sReq += (DeptTimeUTCInMin + DeptTimeZoneInMin + 24*60) % (24*60);
      sReq += ", ";

      // DARP        Char(3)        Аэропорт вылета
      sReq += "'" + DeptArp + "'" + ", ";
      // AARP        Char(3)        Аэропорт прилета
      sReq += "'" + ArvlArp + "'" + ", ";

      // AEFFDATE_L    DateTime    Дни навигации LOC (дата прилета)
      FlightDurationInMin = ArvlTimeUTCInMin - DeptTimeUTCInMin;
      if (FlightDurationInMin <= 0) FlightDurationInMin += (24*60);
      A_EffDateLoc = new Date(EffDate);
      A_EffDateLoc.setMinutes(DeptTimeUTCInMin + FlightDurationInMin + ArvlTimeZoneInMin);
      sReq += "DATETIME(" + A_EffDateLoc.getFullYear() + "," + (A_EffDateLoc.getMonth()+1) + "," + A_EffDateLoc.getDate() + ")";
      sReq += ", ";
      // ADISDATE_L    DateTime    Дни навигации LOC (дата прилета)
      A_DisDateLoc = new Date(DisDate);
      A_DisDateLoc.setMinutes(DeptTimeUTCInMin + FlightDurationInMin + ArvlTimeZoneInMin);
      sReq += "DATETIME(" + A_DisDateLoc.getFullYear() + "," + (A_DisDateLoc.getMonth()+1) + "," + A_DisDateLoc.getDate() + ")";
      sReq += ", ";
      // AFREQ_L    Char(7)        Дни недели прилета LOC
      sReq += "'" + ShiftFreq(Freq, Math.floor((DeptTimeUTCInMin + FlightDurationInMin + ArvlTimeZoneInMin)/(24*60))) + "'";
      sReq += ", ";
      // ATIME_L    Integer        Время прилета LOC (в минутах)
      sReq += (ArvlTimeUTCInMin + ArvlTimeZoneInMin + 24*60) % (24*60);
      sReq += ", ";

      // AIRCTYPE    Char(3)        Тип ВС
      sReq += "'" + AircraftType + "'";
      sReq += ")";
      oConn.Execute(sReq);     
    }
  }
  oConn.Close();
  SSIM.Close();
}

function InitFields()
{
  /*  1 */ fieldStaffId = 0;
  /*  2 */ fieldName = "";
  /*  3 */ fieldPositionA = "";
  /*  4 */ fieldLanguage = "";
  /*  5 */ fieldFlight = 0;
  /*  6 */ fieldPositionN = "";
  /*  7 */ fieldDeptDateTimeUTC = 0;
  /*  8 */ fieldArvlDateTimeUTC = 0;
  /*  9 */ fieldDeptArp = "";
  /* 10 */ fieldArvlArp = "";
  /* 11 */ fieldEmail = "";
  /* 12 */ fieldContact = "";
  /* 13 */ fieldNationality = "";
  /* 14 */ fieldAircraftType = "";
  /* 15 */ fieldAircraftRegNum = "";
 
  sManifestFName = "";
  sManifestFName2 = "";
}

function trim(s)
{
  var re = /^\s+|\s+$/g;
  var str = new String(s);
  return str.replace(re,"");
}

function SeparateStrByChar(s, sep)
{
  var n;
  var ss = "";
 
  if (s.length > 0) ss += s.charAt(0);
  for (n=1; n<s.length; ss += sep + s.charAt(n++));
  return(ss);
}

function SeparateStrByCharExcludeI(s, sep)
{
  var n;
  var ss = "";
 
  if (s.length > 0) ss += s.charAt(0);
  for (n=1; n<s.length; n++)
    if (s.charAt(n) != "I") ss += sep + s.charAt(n);
  return(ss);
}

// Возвращает десятичное число из строки, удаляя начальные нули, чтобы строка не была интерпретирована как восьмеричное число
function StrToInt(s)
{
  var re = /^([\+\-]*)0?([1-9]\d*)/;
  var str = new String(s);
  return parseInt(str.replace(re,"$1$2"));
}

function AircraftTypeMatch(rostset, atype)
{
  var rstr = new String(rostset);
  var s = new String(atype);
  rstr.toUpperCase();
 
  if (rstr.charAt(0)=="X") s = "^"+s;
  s = "["+s+"]";
  var reg = new RegExp(s);
  if (rstr.search(reg) < 0) return(0); else return(1);
}

function ChiefPriority(s)
{
  var str = new String(trim(s));
  switch (str.toUpperCase()) {
    case "P": return (3);
    case "V": return (2);
    case "I": return (1);
    default: return (0);
  }
}

function StrPosA(s)
{
  var str = new String(trim(s));
  switch (str.toUpperCase()) {
    case "I": return ("ИПБ");
    case "D": return ("ГМ");
    case "P": return ("шсб");
    case "V": return ("вшсб");
    default: return ("БП");
  }
}

function DTFormatForManifest(dt)
{
  var s;
  s = dt.getFullYear() + "-";
  if (dt.getMonth() < 9) s += "0";
  s += (dt.getMonth()+1) + "-";
  if (dt.getDate() < 10) s += "0";
  s += dt.getDate() + " ";
  if (dt.getHours() < 10) s += "0";
  s += dt.getHours() + ":";
  if (dt.getMinutes() < 10) s += "0";
  s += dt.getMinutes() + ":";
  if (dt.getSeconds() < 10) s += "0";
  s += dt.getSeconds();
  return(s);
}

function AFLAircraftTypeToIATA(afl)
{
  switch(trim(afl)) {
    case "0": return("319");
    case "1": return("333");
    case "2": return("321");
    case "3": return("SU9");
    case "4": return("738");
    case "6": return("320");
    case "7": return("763");
    case "8": return("77W");
    case "9": return("I93");
    default:  return("");
  }
}

function AppendToManifestFull(arrListOfCrew)
{
  var FSO;
  var ForAppending = 8;
  var TristateTrue = -1; // Unicode mode
  var CM;
  var textCM;
  var i;

  var sFileName = "CrewManifestFull_[";
  sFileName += fieldDeptDateTimeUTC.getFullYear() + "-";
  if (fieldDeptDateTimeUTC.getMonth() < 9) sFileName += "0";
  sFileName += (fieldDeptDateTimeUTC.getMonth()+1) + "-";
  if (fieldDeptDateTimeUTC.getDate() < 10) sFileName += "0";
  sFileName += fieldDeptDateTimeUTC.getDate() + "'T'";
  if (fieldDeptDateTimeUTC.getHours() < 10) sFileName += "0";
  sFileName += fieldDeptDateTimeUTC.getHours();
  if (fieldDeptDateTimeUTC.getMinutes() < 10) sFileName += "0";
  sFileName += fieldDeptDateTimeUTC.getMinutes();
  if (fieldDeptDateTimeUTC.getSeconds() < 10) sFileName += "0";
  sFileName += fieldDeptDateTimeUTC.getSeconds() + "Z].csv";
 
  // Создать поддиректорию и файл, если отсутствуют
  FSO = WScript.CreateObject("Scripting.FileSystemObject");
  if (!FSO.FolderExists(sManifestPath))
    FSO.CreateFolder(sManifestPath);
  if (!FSO.FileExists(sManifestPath+"\\"+sFileName))
    FSO.CreateTextFile(sManifestPath+"\\"+sFileName);
//  CM = FSO.OpenTextFile(sManifestPath+"\\"+sFileName, ForAppending, true, TristateTrue);

  CM = WScript.CreateObject("ADODB.Stream");
  CM.Charset = "UTF-8";
  CM.Type = 2;    // adTypeText
  CM.Mode = 3;    // adModeReadWrite
  CM.Open();
  CM.LoadFromFile(sManifestPath+"\\"+sFileName);
  textCM = "" + CM.ReadText();
  CM.Position = 0;
 
  for (i=0; i<arrListOfCrew.length; i++) {
    textCM += arrListOfCrew[i][indexStaffId] + "|";
    textCM += arrListOfCrew[i][indexName] + "|";
    textCM += arrListOfCrew[i][indexPositionA] + "|";
    textCM += SeparateStrByCharExcludeI(trim(arrListOfCrew[i][indexLanguage]),",") + "|";
    textCM += fieldFlight + "|";
    textCM += arrListOfCrew[i][indexPositionN] + "|";
    textCM += DTFormatForManifest(fieldDeptDateTimeUTC) + "|";
    textCM += DTFormatForManifest(fieldArvlDateTimeUTC) + "|";
    textCM += fieldDeptArp + "|";
    textCM += fieldArvlArp + "|";
    textCM += arrListOfCrew[i][indexEmail] + "|";
    textCM += arrListOfCrew[i][indexContact] + "|";
    textCM += arrListOfCrew[i][indexNationality] + "|";
//    textCM += fieldAircraftType + "|";
    textCM += arrListOfCrew[i][indexAircraftType] + "|";
    textCM += fieldAircraftRegNum + "\r\n";   
   
//    CM.WriteLine(sLine);
  }
  CM.WriteText(textCM);
  CM.SaveToFile(sManifestPath+"\\"+sFileName, 2); // adSaveCreateOverWrite
  CM.Close();
}

function GetBrigadaList(nTrip, dtStartL)
{
  var rsBrigada;
  var sOut = "";
  var sReqBrigada;
  var iMaxPriority;
  var arrAttendant, arrListOfCrew;
  var indexCrew;

  rsBrigada = WScript.CreateObject("ADODB.Recordset");
  rsBrigada.CursorType = 3;

  sReqBrigada = "SELECT REIS.TAB, PERSONAL.FAM, PERSONAL.NAM, PERSONAL.SURNAM, PERSONAL.TYPE AS PERMITTED_TYPES, PLAN_NAR.TYPE AS PLANE_TYPE, PERSONAL.LANG, REIS.N_TRIP, REIS.DATE_BEG, PERSONAL.N_BRIG, REIS.FLAG, PERSONAL.BRIGADIR, PERSONAL.SPEC ";
  sReqBrigada += "FROM (REIS INNER JOIN PERSONAL ON REIS.TAB = PERSONAL.TAB) INNER JOIN PLAN_NAR ON (REIS.DATE_BEG = PLAN_NAR.DATE_BEG) AND (REIS.N_TRIP = PLAN_NAR.N_TRIP) ";
  sReqBrigada += "WHERE (((REIS.N_TRIP)='";
  if (nTrip<10) sReqBrigada += "0";  // Догоняем номер рейса до 3 символов (REIS.N_TRIP имеет символьный тип!)
  if (nTrip<100) sReqBrigada += "0";
  sReqBrigada += nTrip+"') AND ((REIS.DATE_BEG)=DATE(";
  sReqBrigada += dtStartL.getFullYear() + ",";
  sReqBrigada += (dtStartL.getMonth()+1) + ",";
  sReqBrigada += dtStartL.getDate();
  sReqBrigada += ")) AND ((PERSONAL.N_BRIG)<9000) AND (ASC(REIS.FLAG)<>95));";
  rsBrigada.Open(sReqBrigada, sConnect);
  WScript.Echo("  Crew consists of " + rsBrigada.RecordCount + " attendants:");
 
  arrListOfCrew = new Array(rsBrigada.RecordCount);
  indexCrew = 0;

// Первый проход: вычисление полей манифеста
  iMaxPriority = 0; 
  while (!rsBrigada.EOF) {
    arrAttendant = new Array();
    /* 1. fieldStaffId */
    arrAttendant[indexStaffId] = StrToInt(rsBrigada.Fields("TAB"));
    /* 2. fieldName */
    arrAttendant[indexName] = trim(rsBrigada.Fields("FAM")) + " " + trim(rsBrigada.Fields("NAM")) + " " + trim(rsBrigada.Fields("SURNAM"));
   
    // Вычисление приоритетов и поиск максимума
    arrAttendant[indexChiefPriority] = ChiefPriority(rsBrigada.Fields("SPEC"))*AircraftTypeMatch(rsBrigada.Fields("BRIGADIR"), rsBrigada.Fields("PLANE_TYPE"));
    if (arrAttendant[indexChiefPriority] > iMaxPriority) iMaxPriority = arrAttendant[indexChiefPriority];
   
    /* 3. fieldPositionA */
    if (AircraftTypeMatch(rsBrigada.Fields("PERMITTED_TYPES"), rsBrigada.Fields("PLANE_TYPE")) == 1)
      arrAttendant[indexPositionA] = StrPosA(rsBrigada.Fields("SPEC"));
    else
      arrAttendant[indexPositionA] = "ст";

    /* 4. fieldLanguage */
    arrAttendant[indexLanguage] = new String(rsBrigada.Fields("LANG"));
    /* 6. fieldPositionN не используется */
    arrAttendant[indexPositionN] = "";
    /* 11. fieldEmail */
    arrAttendant[indexEmail] = "";
    /* 12. fieldContact */
    arrAttendant[indexContact] = "";
    /* 13. fieldNationality */
    arrAttendant[indexNationality] = "";
    /* 14. fieldAircraftType */
    arrAttendant[indexAircraftType] = new String(AFLAircraftTypeToIATA(rsBrigada.Fields("PLANE_TYPE")));
//    fieldAircraftType = new String(AFLAircraftTypeToIATA(rsBrigada.Fields("PLANE_TYPE")));
   
    arrListOfCrew[indexCrew++] = arrAttendant;
    rsBrigada.MoveNext();
  }

  // Второй проход: назначение бригадира (СБ)
  for (indexCrew=0; indexCrew<rsBrigada.RecordCount; indexCrew++) {
    if (arrListOfCrew[indexCrew][indexChiefPriority] == iMaxPriority)
        arrListOfCrew[indexCrew][indexPositionA] = (arrListOfCrew[indexCrew][indexPositionA] == "ИПБ") ? "СБ/ИПБ" : "СБ";
       
    sOut = "    " + arrListOfCrew[indexCrew][indexStaffId];
    sOut += " " + arrListOfCrew[indexCrew][indexPositionA];
    sOut += "\t"+arrListOfCrew[indexCrew][indexName];
    WScript.Echo(sOut); 
  }
  rsBrigada.Close();
  return(arrListOfCrew);
}

function RunFlightList(dtStartDateLoc)
{
  var rsGenFlightList, rsOut, rsIn;
  var sOut;
  var sRequest;
  var sTrip;
  var arrFlight;
  var iFlightOut, iFlightIn;
  var dtStartLoc = new Date(dtStartDateLoc);
  var dtFinishLoc = new Date();
  var sVRout, sVRin;
  var sReqOutFlight, sReqInFlight;
  var wod;
  var deltaT, dTimeL;
  var arrBrigada;
  var duration;

  rsGenFlightList = WScript.CreateObject("ADODB.Recordset");
  rsGenFlightList.CursorType = 3;

  sRequest = "SELECT MIN(TESTERS.TAB_N) AS S_TAB_N, REIS.DATE_BEG, REIS.DATE_END, REIS.N_TRIP, PLAN_NAR.N_FLIGHT, PLAN_NAR.VR_OUT, PLAN_NAR.VR_IN, REIS.FLAG, REIS.CITY ";
  sRequest += "FROM (TESTERS INNER JOIN REIS ON TESTERS.TAB_N = REIS.TAB) INNER JOIN PLAN_NAR ON (REIS.DATE_BEG = PLAN_NAR.DATE_BEG) AND (REIS.N_TRIP = PLAN_NAR.N_TRIP) ";
  sRequest += "GROUP BY REIS.DATE_BEG, REIS.DATE_END, REIS.N_TRIP, PLAN_NAR.N_FLIGHT, PLAN_NAR.VR_OUT, PLAN_NAR.VR_IN, REIS.FLAG, REIS.CITY ";
  sRequest += "HAVING ((REIS.DATE_BEG = DATE("; // {^2014-07-30}
  sRequest += dtStartLoc.getFullYear() + ",";
  sRequest += (dtStartLoc.getMonth()+1) + ",";
  sRequest += dtStartLoc.getDate();
  sRequest += ")) AND (ASC(REIS.FLAG)<>95) AND (NOT (EMPTY(PLAN_NAR.N_FLIGHT))))";

  rsGenFlightList.Open(sRequest, sConnect);
  sOut = "Found " + rsGenFlightList.RecordCount + " flight";
  if (rsGenFlightList.RecordCount!=1) sOut+="s";
  sOut += " for " + dtStartLoc.toDateString() + "." + "\n";
  WScript.Echo(sOut);

  while (!rsGenFlightList.EOF) {
    InitFields();
    sTrip = new String(rsGenFlightList.Fields("N_FLIGHT"));
    arrFlight = sTrip.split(/\s+/);
    iFlightOut = StrToInt(arrFlight[0]);
    iFlightIn = StrToInt(arrFlight[1]);

    sVRout = new String(rsGenFlightList.Fields("VR_OUT"));
    dtStartLoc.setHours(sVRout.substr(0,2), sVRout.substr(2,2), 0);
    sVRin = new String(rsGenFlightList.Fields("VR_IN"));
    dtFinishLoc = new Date(rsGenFlightList.Fields("DATE_END"));
    dtFinishLoc.setHours(sVRin.substr(0,2), sVRin.substr(2,2), 0);
       
    WScript.Echo("---------------------------------\n" + "SU" + StrToInt(rsGenFlightList.Fields("N_TRIP")));
    arrBrigada = GetBrigadaList(StrToInt(rsGenFlightList.Fields("N_TRIP")), dtStartLoc);   
   
    // Запрос на данные по прямому рейсу
    wod = dtStartLoc.getDay();
    if (wod == 0) wod = 7;
    sReqOutFlight = "SELECT * FROM TIMETABLE ";
    sReqOutFlight += "WHERE ((TIMETABLE.FLIGHTNUM=" + iFlightOut + ") AND "; // TIMETABLE.FLIGHTNUM имеет целочисленный тип!
    sReqOutFlight += "(DATE(";
    sReqOutFlight += dtStartLoc.getFullYear() + ",";
    sReqOutFlight += (dtStartLoc.getMonth()+1) + ",";
    sReqOutFlight += dtStartLoc.getDate();
    sReqOutFlight += ") BETWEEN TIMETABLE.EFFDATE_L AND TIMETABLE.DISDATE_L) AND ";
    sReqOutFlight += "('" + wod + "' $ TIMETABLE.FREQ_L))";

    rsOut = WScript.CreateObject("ADODB.Recordset");
    rsOut.CursorType = 3;
    rsOut.Open(sReqOutFlight, sConnect);
    sOut = "  Start  " + "SU" + arrFlight[0] + ". " + "Matches: " + rsOut.RecordCount + "." + "\n";
   
    if (rsOut.RecordCount>0) {
      /* 5. fieldFlight */
      fieldFlight = iFlightOut;
      deltaT = 1440;
      dTimeL = 1440;
      do {
        if (Math.abs(rsOut.Fields("DTIME_L")-(dtStartLoc.getHours()*60+dtStartLoc.getMinutes())) < deltaT) {
          fieldDeptDateTimeUTC = new Date(dtStartLoc);
          fieldDeptDateTimeUTC.setMinutes(fieldDeptDateTimeUTC.getMinutes() - parseInt(rsOut.Fields("DTIMEZONE")));
          duration = parseInt(rsOut.Fields("ATIME_U")) - parseInt(rsOut.Fields("DTIME_U"));
          if (duration <= 0) duration += (24*60);
          fieldArvlDateTimeUTC = new Date(fieldDeptDateTimeUTC);
          fieldArvlDateTimeUTC.setMinutes(fieldArvlDateTimeUTC.getMinutes() + duration);
          fieldDeptArp = new String(rsOut.Fields("DARP"));
          fieldArvlArp = new String(rsOut.Fields("AARP"));
//          fieldAircraftType = new String(rsOut.Fields("AIRCTYPE"));
          dTimeL = parseInt(rsOut.Fields("DTIME_L"));
          deltaT = Math.abs(rsOut.Fields("DTIME_L")-(dtStartLoc.getHours()*60+dtStartLoc.getMinutes()));
        }
        rsOut.MoveNext();
      } while (!rsOut.EOF);
      sOut += "    Use for start loctime  ";
      if (Math.floor(dTimeL/60)<10) sOut += "0";
      sOut += Math.floor(dTimeL/60) + ":";
      if (dTimeL%60 < 10) sOut += "0";
      sOut += dTimeL % 60 + ", ";
      sOut += fieldDeptArp + "-" + fieldArvlArp + ", aircraft " + arrBrigada[0][indexAircraftType] + "\n";
     
      AppendToManifestFull(arrBrigada);
    }
    else
      sOut += "    NO MANIFEST WAS CREATED!\n";
    rsOut.Close();

    // Запрос на данные по обратному рейсу
    wod = dtFinishLoc.getDay();
    if (wod == 0) wod = 7;
    sReqInFlight = "SELECT * FROM TIMETABLE ";
    sReqInFlight += "WHERE ((TIMETABLE.FLIGHTNUM=" + iFlightIn + ") AND "; // TIMETABLE.FLIGHTNUM имеет целочисленный тип!
    sReqInFlight += "(DATE(";
    sReqInFlight += dtFinishLoc.getFullYear() + ",";
    sReqInFlight += (dtFinishLoc.getMonth()+1) + ",";
    sReqInFlight += dtFinishLoc.getDate();
    sReqInFlight += ") BETWEEN TIMETABLE.AEFFDATE_L AND TIMETABLE.ADISDATE_L) AND ";
    sReqInFlight += "('" + wod + "' $ TIMETABLE.AFREQ_L))";
   
    rsIn = WScript.CreateObject("ADODB.Recordset");
    rsIn.CursorType = 3;
    rsIn.Open(sReqInFlight, sConnect);
    sOut += "  Finish " + "SU"+arrFlight[1]+". " + "Matches: " + rsIn.RecordCount + "." + "\n";   
   
    if (rsIn.RecordCount>0) {
      /* 5. fieldFlight */
      fieldFlight = iFlightIn;
      deltaT = 1440;
      dTimeL = 1440;
      do {
        if (Math.abs(rsIn.Fields("ATIME_L")-(dtFinishLoc.getHours()*60+dtFinishLoc.getMinutes())) < deltaT) {
          fieldArvlDateTimeUTC = new Date(dtFinishLoc);
          fieldArvlDateTimeUTC.setMinutes(fieldArvlDateTimeUTC.getMinutes() - parseInt(rsIn.Fields("ATIMEZONE")));
          duration = parseInt(rsIn.Fields("ATIME_U")) - parseInt(rsIn.Fields("DTIME_U"));
          if (duration <= 0) duration += (24*60);
          fieldDeptDateTimeUTC = new Date(fieldArvlDateTimeUTC);
          fieldDeptDateTimeUTC.setMinutes(fieldDeptDateTimeUTC.getMinutes() - duration);
          fieldDeptArp = new String(rsIn.Fields("DARP"));
          fieldArvlArp = new String(rsIn.Fields("AARP"));
//          fieldAircraftType = new String(rsIn.Fields("AIRCTYPE"));
          dTimeL = parseInt(rsIn.Fields("ATIME_L"));
          deltaT = Math.abs(rsIn.Fields("ATIME_L")-(dtFinishLoc.getHours()*60+dtFinishLoc.getMinutes()));
        }
        rsIn.MoveNext();
      } while (!rsIn.EOF);
      sOut += "    Use for finish loctime ";
      if (Math.floor(dTimeL/60)<10) sOut += "0";
      sOut += Math.floor(dTimeL/60) + ":";
      if (dTimeL%60 < 10) sOut += "0";
      sOut += dTimeL % 60 + ", ";
      sOut += fieldDeptArp + "-" + fieldArvlArp + ", aircraft " + arrBrigada[0][indexAircraftType] + "\n";
     
      AppendToManifestFull(arrBrigada);
    }
    else
      sOut += "    NO MANIFEST WAS CREATED!\n";
    rsIn.Close();

    WScript.Echo(sOut);
    rsGenFlightList.MoveNext();
  }

  rsGenFlightList.Close();
}

function PrintVersionInfo()
{
   var s;
   s = ScriptEngine() + " Version ";
   s += ScriptEngineMajorVersion() + ".";
   s += ScriptEngineMinorVersion() + ".";
   s += ScriptEngineBuildVersion();
   WScript.Echo(s);
}

function Main()
{
  var currDate;
  var fso = WScript.CreateObject("Scripting.FileSystemObject");

    if (WScript.Arguments.Named.Exists("V")) {
    WScript.Echo("Run " + WScript.ScriptFullName + " at " + startTime + ".");
    PrintVersionInfo();
  }
  if (!fso.FolderExists(sDBSourcePath))
    fso.CreateFolder(sDBSourcePath);
  if (WScript.Arguments.Named.Exists("N")) {
    if (WScript.Arguments.Named.Exists("V")) WScript.Echo("Copying table PERSONAL...");
    fso.CopyFile("\\\\MLK-APP-TRUCK01\\Crew_Plan\\CREW1\\OTDEL\\PERSONAL.CDX",sDBSourcePath+"\\",true);
    fso.CopyFile("\\\\MLK-APP-TRUCK01\\Crew_Plan\\CREW1\\OTDEL\\PERSONAL.DBF",sDBSourcePath+"\\",true);
    if (WScript.Arguments.Named.Exists("V")) WScript.Echo("...success.");
    if (WScript.Arguments.Named.Exists("V")) WScript.Echo("Copying table PLAN_NAR...");
    fso.CopyFile("\\\\MLK-APP-TRUCK01\\Crew_Plan\\CREW1\\PLAN_NAR\\PLAN_NAR.CDX",sDBSourcePath+"\\",true);
    fso.CopyFile("\\\\MLK-APP-TRUCK01\\Crew_Plan\\CREW1\\PLAN_NAR\\PLAN_NAR.DBF",sDBSourcePath+"\\",true);
    if (WScript.Arguments.Named.Exists("V")) WScript.Echo("...success.");
    if (WScript.Arguments.Named.Exists("V")) WScript.Echo("Copying table REIS...");
    fso.CopyFile("\\\\MLK-APP-TRUCK01\\Crew_Plan\\CREW1\\WHAT_DO\\REIS.CDX",sDBSourcePath+"\\",true);
    fso.CopyFile("\\\\MLK-APP-TRUCK01\\Crew_Plan\\CREW1\\WHAT_DO\\REIS.DBF",sDBSourcePath+"\\",true);
    if (WScript.Arguments.Named.Exists("V")) WScript.Echo("...success.");
 }
  if (WScript.Arguments.Named.Exists("S")) {
    // Здесь надо скачать и распаковать http://mlk-web-dusid.msk.aeroflot.ru/Schedule/IATA_Current/Current_SSIM.rar !!!!!!!
    ReadSSIM(sDBSourcePath+"\\"+sSSIMFileName);
    WScript.Echo("SSIM was read successfully.");
  }
  if (WScript.Arguments.Named.Exists("D"))
    currDate = new Date(WScript.Arguments.Named("D"))
  else {
    currDate = new Date();
    var dayOfMonth = currDate.getDate();
    currDate.setDate(dayOfMonth + 1); // Tomorrow
  }
  if (WScript.Arguments.Named.Exists("?")) {
    WScript.Echo("Script creates or supplements CrewManifest files from TRUCK rostering database.\n");
    WScript.Echo("makecm.js [/D:yyyy/mm/dd] [/S] [/T] [/SU:n [/1]] [/N] [/V]\n");
    WScript.Echo("/D:yyyy/mm/dd \tDate (local time) of flights.");
    WScript.Echo("\t\tTomorrow's date will be used if /D is not specified.");
    WScript.Echo("/S \t\tRead SSIM and renew TIMETABLE.DBF");
    WScript.Echo("/T \t\tFind the flights only for testers group defined in TESTERS.DBF.");
    WScript.Echo("/SU:n \t\tFlights number of SU (pair of flights in one trip).");
    WScript.Echo("/1 \t\tOnly one flight specified by number in /SU? not pair.");
    WScript.Echo("/N \t\tRenew local copy of tables from \\MLK-APP-TRUCK01\Crew_Plan");
    WScript.Echo("/V \t\tOutput verbose info during the runtime.");
    WScript.Echo("/? \t\tShow this help.");
  }
  else
    RunFlightList(currDate);
}

var startTime, endTime;
startTime = new Date();
Main();
endTime = new Date();
if (WScript.Arguments.Named.Exists("V")) WScript.Echo("Total runtime: ", (endTime.getTime()-startTime.getTime())/1000, " seconds.");