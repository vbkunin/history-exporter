<?php
/**
 * Sample extension to show how adding menu items in iTop
 * This extension does nothing really useful but shows how to use the three possible
 * types of menu items:
 *
 * - An URL to any web page
 * - A Javascript function call
 * - A separator (horizontal line in the menu)
 */
class ExportHistoryMenuItem implements iPopupMenuExtension
{
	/**
	 * Get the list of items to be added to a menu.
	 *
	 * This method is called by the framework for each menu.
	 * The items will be inserted in the menu in the order of the returned array.
	 * @param int $iMenuId The identifier of the type of menu, as listed by the constants MENU_xxx
	 * @param mixed $param Depends on $iMenuId, see the constants defined above
	 * @return object[] An array of ApplicationPopupMenuItem or an empty array if no action is to be added to the menu
	 */
	public static function EnumItems($iMenuId, $param)
	{
		$aResult = array();
		if (self::IsEnabled($param) && $iMenuId == iPopupMenuExtension::MENU_OBJDETAILS_ACTIONS)
		{
			// Разделитель
			$aResult[] = new SeparatorPopupMenuItem(); // Note: separator does not work in iTop 2.0 due to Trac #698, fixed in 2.0.1

			$oObj = $param;

			// Стантартный фильтр истории
			$oFilter = new DBObjectSearch('CMDBChangeOp');
			$oFilter->AddCondition('objkey', $oObj->GetKey(), '=');
			$oFilter->AddCondition('objclass', get_class($oObj), '=');


			// Такой фильтр позволяет использовать поля CMDBChange с типом date для Excel
			// $iObjId = $oObj->GetKey();
			// $sObjClass = get_class($oObj);
			// $oFilter = DBobjectSearch::FromOQL("SELECT CMDBChangeOp, CMDBChange FROM CMDBChangeOp JOIN CMDBChange ON CMDBChangeOp.change = CMDBChange.id WHERE CMDBChangeOp.objkey=$iObjId AND CMDBChangeOp.objclass='$sObjClass'");

			$sFilter = $oFilter->serialize();

			// $sUrl = ApplicationContext::MakeObjectUrl(get_class($oObj), $oObj->GetKey());
			// $sUIPage = cmdbAbstractObject::ComputeStandardUIPage(get_class($oObj));
			// $oAppContext = new ApplicationContext();
			// $sContext = $oAppContext->GetForLink();
			// $oPage->add_linked_script(utils::GetAbsoluteUrlAppRoot().'js/xlsx-export.js');


			// $sExportUrl = utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/export.php';
			// $aResult[] = new URLPopupMenuItem('HistoryExporter:Menu', Dict::S('HistoryExporter:Menu'), $sExportUrl."?filter=".urlencode($sFilter), '_blank');

			// Add a new menu item that triggers a custom JS function defined in our own javascript file: js/sample.js
			$sJSFilter = addslashes($sFilter);
			$sModuleDir = basename(dirname(__FILE__));
			$sJSFileUrl = utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/js/history-export.js';
			$sAjaxFileUrl = utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/ajax.render.php';
			$aResult[] = new JSPopupMenuItem('HistoryExporter:Menu', Dict::S('HistoryExporter:Menu'), "HistoryExportDialog('$sJSFilter', '$sAjaxFileUrl');", array($sJSFileUrl));
		}
		return $aResult;
	}

	/**
	 * Проверяет, включен ли модуль для данного класса
	 */
	protected static function IsEnabled($oObject)
	{
		$sModuleName = basename(dirname(__FILE__)); // Имя модуля должно совпадать с именем каталога
		$param = utils::GetConfig()->GetModuleSetting($sModuleName, 'enabled');
		if (gettype($param) == 'boolean')
		{
			return $param;
		}
		elseif (gettype($param) == 'string')
		{
			$aClasses = array_map('trim', explode(",", $param));
			foreach ($aClasses as $sClass)
			{
				if ($oObject instanceof $sClass)
				{
					return true;
				}
			}
			return false;
		}
		else
		{
			return false;
		}
	}
}