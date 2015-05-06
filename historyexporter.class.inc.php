<?php

require_once(APPROOT.'/application/excelexporter.class.inc.php');

/**
*
*/
class HistoryExporter extends ExcelExporter
{

  // function __construct()
  // {
  //   return parent::__construct();
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
          $oIDSet = new DBObjectSet($this->oSearch, array('date'=>false));
          // $oIDSet->OptimizeColumnLoad(array('id'));
          $oIDSet->Rewind();
          $this->aObjectsIDs = array();
          while($oObj = $oIDSet->Fetch())
          {
            $this->aObjectsIDs[] = $oObj->GetKey();
          }
          $sCode = 'retrieving-data';
          $iPercentage = 5;
          $sMessage = Dict::S('ExcelExporter:RetrievingData');
          $this->iPosition = 0;
          $this->aStatistics['objects_count'] = count($this->aObjectsIDs);
          $this->aStatistics['data_retrieval_duration'] += microtime(true) - $fTime;

          // The first line of the file is the "headers" specifying the label and the type of each column
          $this->GetFieldsList($oIDSet, $this->bAdvancedMode, true, array('CMDBChangeOp.change', 'CMDBChange.date', 'CMDBChangeOp.userinfo'));
          var_dump($this->aTableHeaders);
          $sRow = json_encode($this->aTableHeaders);
          $hFile = @fopen($this->GetDataFile(), 'ab');
          if ($hFile === false)
          {
            throw new Exception('ExcelExporter: Failed to open temporary data file: "'.$this->GetDataFile().'" for writing.');
          }
          fwrite($hFile, $sRow."\n");
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
          $oSet = new DBObjectSet($oCurrentSearch);
          $this->GetFieldsList($oSet, $this->bAdvancedMode, true, array('CMDBChangeOp.change', 'CMDBChange.date', 'CMDBChangeOp.userinfo'));
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

}

?>