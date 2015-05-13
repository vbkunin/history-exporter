function HistoryExport(sFilter)
{
  var sUrl = GetAbsoluteUrlModulesRoot()+'history-exporter/export.php';
  $.post(sUrl, {filter: sFilter}, function(data) {
    $('body').append(data);
  });
}