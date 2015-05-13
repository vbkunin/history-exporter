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

		switch($iMenuId) // type of menu in which to add menu items
		{
			/**
			 * Insert an item into the Actions menu of a list
			 *
			 * $param is a DBObjectSet containing the list of objects
			 */
			case iPopupMenuExtension::MENU_OBJLIST_ACTIONS:

			// Add a new menu item that triggers a custom JS function defined in our own javascript file: js/sample.js
			// $sModuleDir = basename(dirname(__FILE__));
			// $sJSFileUrl = utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/js/sample.js';
			// $iCount = $param->Count(); // number of objects in the set
			// $aResult[] = new JSPopupMenuItem('_Custom_JS_', 'Custom JS Function On List...', "MyCustomJSListFunction($iCount)", array($sJSFileUrl));
			break;

			/**
			 * Insert an item into the Toolkit menu of a list
			 *
			 * $param is a DBObjectSet containing the list of objects
			 */
			case iPopupMenuExtension::MENU_OBJLIST_TOOLKIT:
			break;

			/**
			 * Insert an item into the Actions menu on an object's details page
			 *
			 * $param is a DBObject instance: the object currently displayed
			 */
			case iPopupMenuExtension::MENU_OBJDETAILS_ACTIONS:
			// For any object, add a menu "Google this..." that opens google search in another window
			// with the name of the object as the text to search
			// $aResult[] = new URLPopupMenuItem('_Google_this_', 'Google this...', "http://www.google.com?q=".$param->GetName(), '_blank');

			if ($param instanceof Ticket)
			{
				// add a separator
				$aResult[] = new SeparatorPopupMenuItem(); // Note: separator does not work in iTop 2.0 due to Trac #698, fixed in 2.0.1

        $oObj = $param;
        $oFilter = DBobjectSearch::FromOQL("SELECT ".get_class($oObj)." WHERE id=".$oObj->GetKey());
        $sFilter = $oFilter->serialize();
        // $sUrl = ApplicationContext::MakeObjectUrl(get_class($oObj), $oObj->GetKey());
        // $sUIPage = cmdbAbstractObject::ComputeStandardUIPage(get_class($oObj));
        // $oAppContext = new ApplicationContext();
        // $sContext = $oAppContext->GetForLink();
        // $oPage->add_linked_script(utils::GetAbsoluteUrlAppRoot().'js/xlsx-export.js');
        $sJSFilter = addslashes($sFilter);

				// Add a new menu item that triggers a custom JS function defined in our own javascript file: js/sample.js
				$sModuleDir = basename(dirname(__FILE__));
				// $sExportUrl = utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/export.php';
				// $aResult[] = new URLPopupMenuItem('UI:Menu:ExportHistory', Dict::S('UI:Menu:ExportHistory'), $sExportUrl."?filter=".urlencode($sFilter), '_blank');
        // $aResult[] = new URLPopupMenuItem('UI:Menu:ExportHistory', Dict::S('UI:Menu:ExportHistory'), utils::GetAbsoluteUrlAppRoot()."pages/$sUIPage?operation=search&filter=".urlencode($sFilter)."&format=csv&{$sContext}"),

				$sJSFileUrl = utils::GetAbsoluteUrlModulesRoot().$sModuleDir.'/js/historyexport.js';
				$aResult[] = new JSPopupMenuItem('UI:Menu:ExportHistory', Dict::S('UI:Menu:ExportHistory'), "HistoryExport('$sJSFilter');", array($sJSFileUrl));

			}
			break;

			/**
			 * Insert an item into the Dashboard menu
			 *
			 * The dashboad menu is shown on the top right corner of the page when
			 * a dashboard is being displayed.
			 *
			 * $param is a Dashboard instance: the dashboard currently displayed
			 */
			case iPopupMenuExtension::MENU_DASHBOARD_ACTIONS:
			break;

			/**
			 * Insert an item into the User menu (upper right corner of the page)
			 *
			 * $param is null
			 */
			case iPopupMenuExtension::MENU_USER_ACTIONS:
			break;

		}
		return $aResult;
	}
}