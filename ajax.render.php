<?php

/**
 * Handles history export ajax requests
 *
 * @copyright   Vladimir Kunin v.b.kunin@gmail.com
 * @license     http://opensource.org/licenses/AGPL-3.0
 */

if (!defined('__DIR__')) define('__DIR__', dirname(__FILE__));
require_once(__DIR__.'/../../approot.inc.php');
require_once(APPROOT.'/application/application.inc.php');
require_once(APPROOT.'/application/ajaxwebpage.class.inc.php');
require_once(__DIR__.'/historyexporter.class.inc.php');

try
{
  require_once(APPROOT.'/application/startup.inc.php');
  require_once(APPROOT.'/application/user.preferences.class.inc.php');

  require_once(APPROOT.'/application/loginwebpage.class.inc.php');
  LoginWebPage::DoLogin(false /* bMustBeAdmin */, true /* IsAllowedToPortalUsers */); // Check user rights and prompt if needed

  $oPage = new ajax_page("");
  $oPage->no_cache();

  $operation = utils::ReadParam('operation', '');
  $sFilter = stripslashes(utils::ReadParam('filter', '', false, 'raw_data'));
  $sEncoding = utils::ReadParam('encoding', 'serialize');
  $sClass = utils::ReadParam('class', 'MissingAjaxParam', false, 'class');
  $sStyle = utils::ReadParam('style', 'list');

  switch($operation)
  {
    case 'history_export_dialog':
    $sFilter = utils::ReadParam('filter', '', false, 'raw_data');
    $oPage->SetContentType('text/html');
    $oPage->add(
<<<EOF
<style>
 .ui-progressbar {
  position: relative;
}
.progress-label {
  position: absolute;
  left: 50%;
  top: 1px;
  font-size: 11pt;
}
.download-form button {
  display:block;
  margin-left: auto;
  margin-right: auto;
  margin-top: 2em;
}
.ui-progressbar-value {
  background: url(../setup/orange-progress.gif);
}
.progress-bar {
  height: 20px;
}
.statistics > div {
  padding-left: 16px;
  cursor: pointer;
  font-size: 10pt;
  background: url(../images/minus.gif) 0 2px no-repeat;
}
.statistics > div.closed {
  padding-left: 16px;
  background: url(../images/plus.gif) 0 2px no-repeat;
}

.statistics .closed .stats-data {
  display: none;
}
.stats-data td {
  padding-right: 5px;
}
</style>
EOF
    );
    $oPage->add('<div id="HistoryExportDlg">');
    $oPage->add('<div class="export-options">');
    // Режим выгрузки history-only пока не делаем
    // $oPage->add('<p><input type="checkbox" id="export-history-olny"/>&nbsp;<label for="export-history-olny">'.Dict::S('HistoryExporter:HistoryOnlyMode').'</label></p>');
    // $oPage->add('<p style="font-size:10pt;margin-left:2em;margin-top:-0.5em;padding-bottom:1em;">'.Dict::S('HistoryExporter:HistoryOnlyMode+').'</p>');
    $oPage->add('<p><input type="checkbox" id="export-auto-download" checked="checked"/>&nbsp;<label for="export-auto-download">'.Dict::S('ExcelExport:AutoDownload').'</label></p>');
    $oPage->add('</div>');
    $oPage->add('<div class="progress"><p class="status-message">'.Dict::S('ExcelExport:PreparingExport').'</p><div class="progress-bar"><div class="progress-label"></div></div></div>');
    $oPage->add('<div class="statistics"><div class="stats-toggle closed">'.Dict::S('ExcelExport:Statistics').'<div class="stats-data"></div></div></div>');
    $oPage->add('</div>');
    $aLabels = array(
      'dialog_title' => Dict::S('HistoryExporter:ExportDialogTitle'),
      'cancel_button' => Dict::S('UI:Button:Cancel'),
      'export_button' => Dict::S('ExcelExporter:ExportButton'),
      'download_button' => Dict::Format('ExcelExporter:DownloadButton', 'history.xlsx'), //TODO: better name for the file (based on the class of the filter??)
    );
    $sJSLabels = json_encode($aLabels);
    $sFilter = addslashes($sFilter);
    $sModuleDir = basename(dirname(__FILE__));
    $sAjaxPageUrl = addslashes(utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/ajax.render.php');
    $oPage->add_ready_script("$('#HistoryExportDlg').historyexporter({filter: '$sFilter', labels: $sJSLabels, ajax_page_url: '$sAjaxPageUrl'});");
    break;

    case 'xlsx_start':
    $sFilter = utils::ReadParam('filter', '', false, 'raw_data');
    // $bAdvanced = (utils::ReadParam('advanced', 'false') == 'true'); // не используется в истории
    $oSearch = DBObjectSearch::unserialize($sFilter);

    $oHistoryExporter = new HistoryExporter();
    $oHistoryExporter->SetObjectList($oSearch);
    // $oHistoryExporter->SetChunkSize(10);
    // $oHistoryExporter->SetAdvancedMode($bAdvanced);
    $sToken = $oHistoryExporter->SaveState();
    $oPage->add(json_encode(array('status' => 'ok', 'token' => $sToken)));
    break;

    case 'xlsx_run':
    $sMemoryLimit = MetaModel::GetConfig()->Get('xlsx_exporter_memory_limit');
    ini_set('memory_limit', $sMemoryLimit);
    ini_set('max_execution_time', max(300, ini_get('max_execution_time'))); // At least 5 minutes

    $sToken = utils::ReadParam('token', '', false, 'raw_data');
    $oHistoryExporter = new HistoryExporter($sToken);
    $aStatus = $oHistoryExporter->Run();
    $aResults = array('status' => $aStatus['code'], 'percentage' =>  $aStatus['percentage'], 'message' =>  $aStatus['message']);
    if ($aStatus['code'] == 'done')
    {
      $aResults['statistics'] = $oHistoryExporter->GetStatistics('html');
    }
    $oPage->add(json_encode($aResults));
    break;

    case 'xlsx_download':
    $sToken = utils::ReadParam('token', '', false, 'raw_data');
    $oPage->SetContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    $oPage->SetContentDisposition('attachment', 'history.xlsx');
    $sFileContent = HistoryExporter::GetExcelFileFromToken($sToken);
    $oPage->add($sFileContent);
    HistoryExporter::CleanupFromToken($sToken);
    break;

    case 'xlsx_abort':
    // Stop & cleanup an export...
    $sToken = utils::ReadParam('token', '', false, 'raw_data');
    HistoryExporter::CleanupFromToken($sToken);
    break;


    default:
    $oPage->p("Invalid query.");
  }

  $oPage->output();
}
catch (Exception $e)
{
  // note: transform to cope with XSS attacks
  echo htmlentities($e->GetMessage(), ENT_QUOTES, 'utf-8');
  echo "<p>Debug trace: <pre>".$e->getTraceAsString()."</pre></p>\n";
  IssueLog::Error($e->getMessage());
}

?>