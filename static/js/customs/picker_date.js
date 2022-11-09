/* ------------------------------------------------------------------------------
*
*  # Date and time pickers
*
*  Demo JS code for picker_date.html page
*
* ---------------------------------------------------------------------------- */

document.addEventListener('DOMContentLoaded', function() {

    var lang = $('input[name="selected-lang"]').val();
    moment.locale(lang);
    
    // Date ranger picker with single date
    $('.daterange-single').daterangepicker({ 
        singleDatePicker: true
    });

    // Single picker
    $.fn.datepicker.dates['ja'] = {
        days: [ "日曜日","月曜日","火曜日","水曜日","木曜日","金曜日","土曜日" ],
        daysShort: [ "日","月","火","水","木","金","土" ],
        daysMin: [ "日","月","火","水","木","金","土" ],
        months: [ "1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月" ],
        monthsShort: [ "1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月" ],
        today: "今日",
        clear: "削除",
        format: "yyyy/mm/dd",
        titleFormat: "yyyy/mm/dd",
        weekStart: 0,
    };

    $('.datepicker-nullable').datepicker({
        language: lang,
    });
    
});
