<?php

require_once(APPROOT.'/application/excelexporter.class.inc.php');

/**
*
*/
class HistoryExporter extends ExcelExporter
{
  // Is history only mode really needed??
  //
  // protected $bHistoryOnlyMode;
  protected $aDetailsFields;

  // function __construct($sToken = null)
  // {
  //   $this->bHistoryOnlyMode = false;
  //   return parent::__construct($sToken);
  // }

  // function SetHistoryOnlyMode($bHistoryOnly)
  // {
  //   $this->bHistoryOnlyMode = $bHistoryOnly;
  // }

  public function Run()
  {
    $sCode = 'error';
    $iPercentage = 100;
    $sMessage = Dict::Format('ExcelExporter:ErrorUnexpected_State', $this->sState);
    $fTime = microtime(true);

    try
    {
      switch($this->sState)
      {
        case 'new':
          // Формируем информацию об объекте и пишем в файл
          $aQueryParams = $this->oSearch->GetInternalParams();
          if (!(UserRights::IsActionAllowed($aQueryParams['objclass'], UR_ACTION_READ) && (UR_ALLOWED_YES || UR_ALLOWED_DEPENDS)))
          {
            // Если пользователь не имеет прав на просмотр объекта, но каким-то чудом добрался до выгрузки истории
            throw new SecurityException('HistoryExporter: You are not allowed to access the object class '.$aQueryParams['objclass'].'.');
          }
          $oObject = MetaModel::GetObject($aQueryParams['objclass'], $aQueryParams['objkey']);

          $this->GetDetailsFields($oObject);
          // var_dump($this->aDetailsFields);

          $hFile = @fopen($this->GetDetailsFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('HistoryExporter: Failed to open temporary details data file: "'.$this->GetDetailsFile().'" for writing.');
          }

          // Первая строка - Класс и Название объекта
          $aRow = array(MetaModel::GetName(get_class($oObject)), $oObject->GetName());
          $sRow = json_encode($aRow);
          fwrite($hFile, $sRow."\n");

          $aRow = array('', '');
          $sRow = json_encode($aRow);
          fwrite($hFile, $sRow."\n");

          foreach ($this->aDetailsFields as $key => $value) {
            if (is_array($value))
            {
              if ($key == 'Lists' )
              {
                // lists пока не обрабатываем
              }
              elseif ($key == 'CaseLogs')
              {
                foreach ($value as $sAttCode) {
                  $sLabel = $oObject->GetLabel($sAttCode);
                  $oCaseLog = $oObject->Get($sAttCode);
                  $sValue = trim(preg_replace('/========== ([^=]+) ============/', '********** $1 ************', $oCaseLog->GetText()));
                  $aRow = array($sLabel.':', $sValue);
                  $sRow = json_encode($aRow);
                  fwrite($hFile, $sRow."\n");
                }
              }
              else
              {
                // fieldset
                // key - заголовок, value - массив полей
                // | Название набора | пустота |
                $aRow = array(Dict::S($key), '');
                $sRow = json_encode($aRow);
                fwrite($hFile, $sRow."\n");
                foreach ($value as $sAttCode) {
                  // | Название свойства: | Значение свойства |
                  $sLabel = $oObject->GetLabel($sAttCode);
                  $sValue = $oObject->GetEditValue($sAttCode);
                  $aRow = array($sLabel.':', $sValue);
                  $sRow = json_encode($aRow);
                  fwrite($hFile, $sRow."\n");
                }
                // Пустая строка после набора полей
                $aRow = array('', '');
                $sRow = json_encode($aRow);
                fwrite($hFile, $sRow."\n");
              }
            }
            else
            {
              // Обрабатываем отдельные поля
              $sAttCode = $value;
              // | Название свойства: | Значение свойства |
              $sLabel = $oObject->GetLabel($sAttCode);
              $sValue = $oObject->GetEditValue($sAttCode);
              $aRow = array($sLabel.':', $sValue);
              $sRow = json_encode($aRow);
              fwrite($hFile, $sRow."\n");
            }
          }
          fclose($hFile);

          // Считаем изменения объекта и подготавливаем заголовки для листа с историей
          $oIDSet = new CMDBObjectSet($this->oSearch);
          $oIDSet->OptimizeColumnLoad(array('id'));
          // $oIDSet->Rewind(); для выборки истории
          $this->aObjectsIDs = array();
          while($oObj = $oIDSet->Fetch())
          {
            $this->aObjectsIDs[] = $oObj->GetKey();
          }
          $sCode = 'retrieving-data';
          $iPercentage = 5;
          $sMessage = Dict::S('ExcelExporter:RetrievingData');
          $this->iPosition = 0;
          $this->aStatistics['objects_count'] = count($this->aObjectsIDs) + 1; // 1 - сам объект истории
          $this->aStatistics['data_retrieval_duration'] += microtime(true) - $fTime;
          // $this->GetFieldsList($oIDSet, $this->bAdvancedMode, true, array('date', 'userinfo', 'change'));
          // $sRow = json_encode($this->aTableHeaders);
          $aHistoryTableHeaders = array(Dict::S('UI:History:Date') => "datetime", Dict::S('UI:History:User') => "string", Dict::S('UI:History:Changes') => "string");
          $hFile = @fopen($this->GetHistoryFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('HistoryExporter: Failed to open temporary history data file: "'.$this->GetHistoryFile().'" for writing.');
          }
          fwrite($hFile, json_encode($aHistoryTableHeaders)."\n");
          fclose($hFile);

          // Next state
          $this->sState = 'retrieving-data';
        break;

        case 'retrieving-data':
          $oCurrentSearch = clone $this->oSearch;
          $aIDs = array_slice($this->aObjectsIDs, $this->iPosition, $this->iChunkSize);

          $oCurrentSearch->AddCondition('id', $aIDs, 'IN');
          $hFile = @fopen($this->GetHistoryFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('ExcelExporter: Failed to open temporary data file: "'.$this->GetHistoryFile().'" for writing.');
          }
          $oSet = new DBObjectSet($oCurrentSearch, array('date'=>false));
          $oSet->Rewind();
          $this->GetFieldsList($oSet, $this->bAdvancedMode);

          $aChanges= array();
          // Группируем изменения, сделанные за одну операцию
          while($oChangeOp = $oSet->Fetch())
          {
            $sChangeDescription = strip_tags($oChangeOp->GetDescription());
            if ($sChangeDescription != '')
            {
              // The change is visible for the current user
              $changeId = $oChangeOp->Get('change');
              $aChanges[$changeId]['date'] = $oChangeOp->Get('date');
              $aChanges[$changeId]['userinfo'] = $oChangeOp->Get('userinfo');
              if (!isset($aChanges[$changeId]['log']))
              {
                $aChanges[$changeId]['log'] = array();
              }
              $aChanges[$changeId]['log'][] = $sChangeDescription;
            }
          }
          foreach($aChanges as $aChange)
          {
            $aRow = array($aChange['date'], $aChange['userinfo'], implode("\n", $aChange['log']));
            $sRow = json_encode($aRow);
            fwrite($hFile, $sRow."\n");
          }
          fclose($hFile);

          if (($this->iPosition + $this->iChunkSize) > count($this->aObjectsIDs))
          {
            // Next state
            $this->sState = 'building-excel';
            $sCode = 'building-excel';
            $iPercentage = 80;
            $sMessage = Dict::S('ExcelExporter:BuildingExcelFile');
          }
          else
          {
            $sCode = 'retrieving-data';
            $this->iPosition += $this->iChunkSize;
            $iPercentage = 5 + round(75 * ($this->iPosition / count($this->aObjectsIDs)));
            $sMessage = Dict::S('ExcelExporter:RetrievingData');
          }
        break;

        case 'building-excel':

          // Лист с деталями объекта
          $hFile = @fopen($this->GetDetailsFile(), 'rb');
          if ($hFile === false)
          {
            throw new Exception('HistoryExporter: Failed to open temporary details file: "'.$this->GetDetailsFile().'" for reading.');
          }
          $sTitle = fgets($hFile);
          $aTitle = json_decode($sTitle, true);
          // $sDetailsSheet = $aTitle[0].' '.$aTitle[1];
          $sDetailsSheet = Dict::S('UI:PropertiesTab');
          $aDetailsHeaders = array($aTitle[0].':' => 'string', $aTitle[1] => 'string');
          $aDetails = array();
          while($sLine = fgets($hFile))
          {
            $aRow = json_decode($sLine);
            $aDetails[] = $aRow;
          }
          fclose($hFile);
          @unlink($this->GetDetailsFile());

          // Лист с историей
          $hFile = @fopen($this->GetHistoryFile(), 'rb');
          if ($hFile === false)
          {
            throw new Exception('ExcelExporter: Failed to open temporary history file: "'.$this->GetHistoryFile().'" for reading.');
          }
          $sHistoryHeaders = fgets($hFile);
          $aHistoryHeaders = json_decode($sHistoryHeaders, true);
          $sHistorySheet = Dict::S('UI:HistoryTab');
          $aHistory = array();
          while($sLine = fgets($hFile))
          {
            $aRow = json_decode($sLine);
            $aHistory[] = $aRow;
          }
          fclose($hFile);
          @unlink($this->GetHistoryFile());


          $fStartExcel = microtime(true);
          $writer = new XLSXWriter();
          $writer->setAuthor(UserRights::GetUserFriendlyName());
          $writer->writeSheet($aDetails, $sDetailsSheet, $aDetailsHeaders);
          $writer->writeSheet($aHistory, $sHistorySheet, $aHistoryHeaders);
          $fExcelTime = microtime(true) - $fStartExcel;
          $this->aStatistics['excel_build_duration'] = $fExcelTime;

          $fTime = microtime(true);
          $writer->writeToFile($this->GetExcelFilePath());
          $fExcelSaveTime = microtime(true) - $fTime;
          $this->aStatistics['excel_write_duration'] = $fExcelSaveTime;

          // Next state
          $this->sState = 'done';
          $sCode = 'done';
          $iPercentage = 100;
          $sMessage = Dict::S('ExcelExporter:Done');
        break;

        case 'done':
          $this->sState = 'done';
          $sCode = 'done';
          $iPercentage = 100;
          $sMessage = Dict::S('ExcelExporter:Done');
        break;
      }
    }
    catch(Exception $e)
    {
      $sCode = 'error';
      $sMessage = $e->getMessage();
    }

    $this->aStatistics['total_duration'] += microtime(true) - $fTime;
    $peak_memory = memory_get_peak_usage(true);
    if ($peak_memory > $this->aStatistics['peak_memory_usage'])
    {
      $this->aStatistics['peak_memory_usage'] = $peak_memory;
    }

    return array(
      'code' => $sCode,
      'message' => $sMessage,
      'percentage' => $iPercentage,
    );
  }

  /**
   * Возвращает адрес файла с описанием объекта
   *
   */
  protected function GetDetailsFile()
  {
    return APPROOT.'data/bulk_export/'.$this->sToken.'.details';
  }

  /**
   * Возвращает адрес файла с историей изменения объекта
   *
   */
  protected function GetHistoryFile()
  {
    return APPROOT.'data/bulk_export/'.$this->sToken.'.history';
  }

  public static function CleanupFromToken($sToken)
  {
    @unlink(APPROOT.'data/bulk_export/'.$sToken.'.details');
    @unlink(APPROOT.'data/bulk_export/'.$sToken.'.history');
    parent::CleanupFromToken($sToken);
  }


  protected function GetDetailsFields($oObject)
  {
    $sClass = get_class($oObject);
    // Получаем все колонки, наборы (fieldset) и поля для отображения деталей объекта
    if (UserRights::IsPortalUser())
    {
      // Если пользователь портала, ограничить список полей на основе PORTAL_CLASSNAME_DETAILS_ZLIST
      $sPortalConfig = APPROOT.'/portal/config-portal.php'; // дополнительный конфиг-файл портала, в котором объявлены специфически константы PORTAL_MYCLASS_DETAILS_ZLIST
      if (file_exists($sPortalConfig))
      {
        require_once($sPortalConfig);
      }
      $sZListConstName = 'PORTAL_'.strtoupper($sClass).'_DETAILS_ZLIST';
      if (defined($sZListConstName))
      {
        $aDetailsZListItems = json_decode(constant($sZListConstName), true);
      }
      else
      {
        throw new Exception("HistoryExporter: Missing portal constant '$sZListConstName'");
      }
      // Формируем структуру: убираем колонки, оставляем наборы и поля и их порядок
      $aDetails = self::ProcessDetailsList($aDetailsZListItems);
      // Добавляем журналы
      $sPublicLogConstName = 'PORTAL_'.strtoupper($sClass).'_PUBLIC_LOG';
      if (defined($sPublicLogConstName))
      {
        $aDetails['CaseLogs'][] = constant($sPublicLogConstName);
      }
      else
      {
        throw new Exception("HistoryExporter: Missing portal constant '$sPublicLogConstName'");
      }
    }
    else
    {
      // Если пользователь - агент, получаем обычный список полей для деталей
      $aDetailsZListItems = MetaModel::GetZListItems($sClass, 'details');
      // Формируем структуру: убираем колонки, оставляем наборы и поля и их порядок
      $aDetails = self::ProcessDetailsList($aDetailsZListItems);
      // Добавляем журналы
      foreach(MetaModel::ListAttributeDefs($sClass) as $sAttCode => $oAttDef)
      {
        if ($oAttDef instanceof AttributeCaseLog)
        {
          $aDetails['CaseLogs'][] = $sAttCode;
        }
      }
      // var_dump('Структура полностью:');
      // var_dump($aDetails);
    }
    //
    // Если у объекта есть жизненный цикл, то некоторые поля нужно скрыть в зависимости от статуса
    $sStateAttCode = MetaModel::GetStateAttributeCode($sClass);
    if ($sStateAttCode)
    {
      $sState = $oObject->Get($sStateAttCode);
      $aDetails = self::RemoveHiddenFields($sClass, $sState, $aDetails);
      // var_dump('Структура видимая:');
      // var_dump($aDetails);
    }

    $this->aDetailsFields = $aDetails;
    // $this->aDetailsFields = self::FillAttributeDefs($sClass, $aDetails);
    // var_dump('Заполненный:');
    // var_dump($aDetails);
  }

  // static function FillAttributeDefs($sClass, $aDetails)
  // {
  //   $aResult = array();
  //   foreach ($aDetails as $key => $value) {
  //     if (is_array($value))
  //     {
  //       $aResult[$key] = self::FillAttributeDefs($sClass, $value);
  //     }
  //     else
  //     {
  //       $aResult[$value] = MetaModel::GetAttributeDef($sClass, $value);
  //     }
  //   }
  //   return $aResult;
  // }

  static function RemoveHiddenFields($sClass, $sState, $aFieldsList)
  {
    $aResult = array();
    foreach($aFieldsList as $key => $value )
    {
      if (is_array($value))
      {
        // Если это массив => набор полей, повторяем для каждого поля
        $sFieldset = $key;
        $aFields = $value;
        $aResult[$sFieldset] = self::RemoveHiddenFields($sClass, $sState, $aFields);
      }
      else
      {
        // Если это отдельное поле
        $sAttCode = $value;
        $iFlags = MetaModel::GetAttributeFlags($sClass, $sState, $sAttCode);
        if (($iFlags & OPT_ATT_HIDDEN) == 0)
        {
          // Если поле отображается в текущем статусе объекта
          $aResult[] = $sAttCode;
        }
      }
    }
    return $aResult;
  }

  static function ProcessDetailsList($aFieldsList)
  {
    $aResult = array();
    $aListFields = array(); // вкладки со связанными объектами
    foreach($aFieldsList as $key => $value)
    {
      if (!is_array($value))
      {
        if (preg_match('/^(.*)_list$/U', $value))
        {
          $aListFields['Lists'][] = $value;
        }
        else
        {
          $aResult[] = $value;
        }
      }
      elseif (preg_match('/^fieldset:(.*)$/U', $key, $matches))
      {
        $aResult[$matches[1]] = self::ProcessDetailsList($value);
      }
      else
      {
        $aResult = array_merge($aResult,self::ProcessDetailsList($value));
      }
    }
    return array_merge($aResult, $aListFields);
  }

}

?>