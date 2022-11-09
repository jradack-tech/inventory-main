/* ------------------------------------------------------------------------------
*
*  # SelectBoxIt selects
*
*  Demo JS code for form_select_box_it.html page
*
* ---------------------------------------------------------------------------- */

document.addEventListener('DOMContentLoaded', function() {


    // Basic examples
    // ------------------------------

    // Basic initialization
    $(".selectbox, .product-type-selectbox").selectBoxIt({
        autoWidth: false
    });

   
    // Other additions
    // ------------------------------

    // Enable popover
    $('[data-toggle="popover"').popover();
    
});
