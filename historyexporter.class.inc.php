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




          $sDetailsSheet = MetaModel::GetName(get_class($oObject)).": ".$oObject->GetName(); // TODO: перенести в building-excel
          // var_dump($sClassName, $oObject->GetName());
          // $oObject->GetLabel($sAttCode) => $oObject->Get($sAttCodeEx)
          //
          $hFile = @fopen($this->GetDetailsFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('HistoryExporter: Failed to open temporary details data file: "'.$this->GetDetailsFile().'" for writing.');
          }
          $oSet = DBObjectSet::FromObject($oObject);
          $this->GetFieldsList($oSet, $this->bAdvancedMode); // это нужно для проверки aAuthorizedClasses и получения aFieldsList
          while($aObjects = $oSet->FetchAssoc())
          {
            // $aRow = array();
            foreach($this->aAuthorizedClasses as $sAlias => $sClassName)
            {
              $oObj = $aObjects[$sAlias];
              // if ($this->bAdvancedMode)
              // {
              //   $aRow[] = $oObj->GetKey();
              // }
              foreach($this->aFieldsList[$sAlias] as $sAttCodeEx => $oAttDef)
              {
                $sLabel = MetaModel::GetLabel($sClassName, $sAttCodeEx);
                $value = $oObj->Get($sAttCodeEx);
                if ($value instanceOf ormCaseLog)
                {
                  // Extract the case log as text and remove the "===" which make Excel think that the cell contains a formula the next time you edit it!
                  $sExcelVal = trim(preg_replace('/========== ([^=]+) ============/', '********** $1 ************', $value->GetText()));
                }
                else
                {
                  $sExcelVal =  $oAttDef->GetEditValue($value, $oObj);
                }
                $aRow = array($sLabel, $sExcelVal);
                $sRow = json_encode($aRow);
                fwrite($hFile, $sRow."\n");
              }
            }
          }
          fclose($hFile);



          // $hFile = @fopen($this->GetDataFile(), 'ab');
          // if ($hFile === false)
          // {
          //   throw new Exception('ExcelExporter: Failed to open temporary data file: "'.$this->GetDataFile().'" for writing.');
          // }
          // // $sHistoryHeader = json_encode($aHistoryHeader);
          // // fwrite($hFile, json_encode($aHistoryHeader)."\n");
          // // $sHistoryTableHeaders = json_encode($aHistoryTableHeaders);
          // fwrite($hFile, json_encode($aHistoryTableHeaders)."\n");
          // fclose($hFile);




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
          $sHistorySheet = Dict::S('UI:HistoryTab'); // TODO: перенести в building-excel
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
          $hFile = @fopen($this->GetDataFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('ExcelExporter: Failed to open temporary data file: "'.$this->GetDataFile().'" for writing.');
          }
          $oSet = new DBObjectSet($oCurrentSearch, array('date'=>false));
          $oSet->Rewind();
          $this->GetFieldsList($oSet, $this->bAdvancedMode);

          // var_dump($this->aFieldsList);
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

          // while($aObjects = $oSet->FetchAssoc())
          // {
          //   $aRow = array();
          //   foreach($this->aAuthorizedClasses as $sAlias => $sClassName)
          //   {
          //     $oObj = $aObjects[$sAlias];
          //     if ($this->bAdvancedMode)
          //     {
          //       $aRow[] = $oObj->GetKey();
          //     }
          //     foreach($this->aFieldsList[$sAlias] as $sAttCodeEx => $oAttDef)
          //     {
          //       $value = $oObj->Get($sAttCodeEx);
          //       // var_dump($value);
          //       if ($value instanceOf ormCaseLog)
          //       {
          //         // Extract the case log as text and remove the "===" which make Excel think that the cell contains a formula the next time you edit it!
          //         $sExcelVal = trim(preg_replace('/========== ([^=]+) ============/', '********** $1 ************', $value->GetText()));
          //       }
          //       else
          //       {
          //         $sExcelVal =  $oAttDef->GetEditValue($value, $oObj);
          //       }
          //       $aRow[] = $sExcelVal;
          //     }
          //   }
          //   $sRow = json_encode($aRow);
          //   fwrite($hFile, $sRow."\n");
          // }
          // fclose($hFile);

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

        case 'retrieving-history':
          // $this->GetFieldsList($oIDSet, $this->bAdvancedMode, true, array('CMDBChangeOp.change', 'CMDBChange.date', 'CMDBChangeOp.userinfo'));

          $oCurrentSearch = clone $this->oSearch;
          $aIDs = array_slice($this->aObjectsIDs, $this->iPosition, $this->iChunkSize);

          $oCurrentSearch->AddCondition('id', $aIDs, 'IN');
          $hFile = @fopen($this->GetDataFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('ExcelExporter: Failed to open temporary data file: "'.$this->GetDataFile().'" for writing.');
          }
          $oSet = new DBObjectSet($oCurrentSearch);
          $this->GetFieldsList($oSet, $this->bAdvancedMode);
          while($aObjects = $oSet->FetchAssoc())
          {
            $aRow = array();
            foreach($this->aAuthorizedClasses as $sAlias => $sClassName)
            {
              $oObj = $aObjects[$sAlias];
              if ($this->bAdvancedMode)
              {
                $aRow[] = $oObj->GetKey();
              }
              foreach($this->aFieldsList[$sAlias] as $sAttCodeEx => $oAttDef)
              {
                $value = $oObj->Get($sAttCodeEx);
                if ($value instanceOf ormCaseLog)
                {
                  // Extract the case log as text and remove the "===" which make Excel think that the cell contains a formula the next time you edit it!
                  $sExcelVal = trim(preg_replace('/========== ([^=]+) ============/', '********** $1 ************', $value->GetText()));
                }
                else
                {
                  $sExcelVal =  $oAttDef->GetEditValue($value, $oObj);
                }
                $aRow[] = $sExcelVal;
              }
            }
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
          $hFile = @fopen($this->GetDataFile(), 'rb');
          if ($hFile === false)
          {
            throw new Exception('ExcelExporter: Failed to open temporary data file: "'.$this->GetDataFile().'" for reading.');
          }
          $sHeaders = fgets($hFile);
          $aHeaders = json_decode($sHeaders, true);

          $aData = array();
          while($sLine = fgets($hFile))
          {
            $aRow = json_decode($sLine);
            $aData[] = $aRow;
          }
          fclose($hFile);
          @unlink($this->GetDataFile());

          $fStartExcel = microtime(true);
          $writer = new XLSXWriter();
          $writer->setAuthor(UserRights::GetUserFriendlyName());
          $writer->writeSheet($aData,'Sheet1', $aHeaders);
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
    // $oAppContext = new ApplicationContext();
    // $sStateAttCode = MetaModel::GetStateAttributeCode(get_class($oObject));
    $aDetails = array();
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
      $sConstName = 'PORTAL_'.strtoupper($sClass).'_DETAILS_ZLIST';
      if (defined($sConstName))
      {
        // var_dump($sConstName);
        $aDetailsList = json_decode(constant($sConstName), true);
      }
      else
      {
        throw new Exception("HistoryExporter: Missing portal constant '$sConstName'");
      }
    }
    else
    {
      // Если пользователь - агент, получаем обычный список полей для деталей
      $aDetailsList = MetaModel::GetZListItems($sClass, 'details');
    }
    var_dump($aDetailsList);

    // Формируем структуру: убираем колонки, оставляем наборы и поля и их порядок
    // | Название набора  |               |
    // | Название поля    | Значение поля |
    $aDetailsStruct = self::ProcessDetailsList($aDetailsList);

    var_dump($aDetailsStruct);
    // Перебираем поля и оставляем только те, которые отображаются в текущем статусе объекта
    // TODO: убрать в отдельную функцию, которая в случае с набором полей вызывает себя еще раз
    foreach($aDetailsStruct as $key => $value )
    {
      if (is_array($value))
      {
        // Если это набор полей
        $sFieldset = $key;
        $aFields = $value;
        foreach ($aFields as $sAttCode) {
          $iFlags = $oObject->GetAttributeFlags($sAttCode);
          // var_dump($sAttCode, $iFlags & OPT_ATT_HIDDEN);
          if (($iFlags & OPT_ATT_HIDDEN) == 0)
          {
            // Если поле отображается в текущем статусе объекта
            $aDetails[$sFieldset][] = $sAttCode;
          }
        }
      }
      else
      {
        // Если это отдельное поле
        $sAttCode = $value;
        $iFlags = $oObject->GetAttributeFlags($sAttCode);
        // var_dump($sAttCode, $iFlags & OPT_ATT_HIDDEN);
        if (($iFlags & OPT_ATT_HIDDEN) == 0)
        {
          // Если поле отображается в текущем статусе объекта
          $aDetails[] = $sAttCode;
        }
      }
    }

    // var_dump($aDetails);

  }

  static function ProcessDetailsList($aList)
  {
    $aResult = array();
    $aListFields = array(); // вкладки со связанными объектами
    foreach($aList as $key => $value)
    {
      if (!is_array($value))
      {
        if (preg_match('/^(.*)_list$/U', $value))
        {
          $aListFields[] = $value;
        }
        else
        {
          $aResult[] = $value;
        }
      }
      elseif (preg_match('/^fieldset:(.*)$/U', $key))
      {
        $aResult[$key] = $value;
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