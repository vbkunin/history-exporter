<?php
//
// iTop module definition file
//

// Git commit info
if (file_exists(dirname(__FILE__).'/.git/HEAD'))
{
	$sDate = date("d.m.y G:i", filemtime(dirname(__FILE__).'/.git/index'));
	$sHead = trim(preg_replace('/ref:/', '', file_get_contents(dirname(__FILE__).'/.git/HEAD'), -1, $iCount));
	if ($iCount === 1)
	{
		$sCommit = '(commit '.file_get_contents(dirname(__FILE__).'/.git/'.$sHead, false, null, 0, 7).' from '.$sDate.')';
	}
	elseif ($iCount === 0)
	{
		$sCommit = '(commit '.substr($sHead, 0, 7).' from '.$sDate.')';
	}
	else
	{
		$sCommit = '(commit info not found)';
	}
}
else
{
	$sCommit = '';
}


SetupWebPage::AddModule(
	__FILE__, // Path to the current file, all other file names are relative to the directory containing this file
	'history-exporter/',
	array(
		// Identification
		//
		'label' => "History Exporter $sCommit",
		'category' => 'business',

		// Setup
		//
		'dependencies' => array(

		),
		'mandatory' => false,
		'visible' => true,

		// Components
		//
		'datamodel' => array(
			'main.history-exporter.php',
			'historyexporter.class.inc.php'
		),
		'webservice' => array(
		),
		'data.struct' => array(
			// add your 'structure' definition XML files here,
		),
		'data.sample' => array(
			// add your sample data XML files here,
		),

		// Documentation
		//
		'doc.manual_setup' => '', // hyperlink to manual setup documentation, if any
		'doc.more_information' => '', // hyperlink to more information, if any

		// Default settings
		//
		'settings' => array(
			// Module specific settings go here, if any
			'enabled' => true,
		),
	)
);


?>
