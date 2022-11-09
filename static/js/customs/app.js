// Reset the form total number in management form
function resetMangementForm(prefix) {
    var numTotalForms = $('table.table-' + prefix + ' tr[class^="formset_row-"]').length;
    $('#id_' + prefix + '-TOTAL_FORMS').val(numTotalForms);
}

document.addEventListener('DOMContentLoaded', function () {

    // Set formset prefixes here
    var product_prefix = 'product';
    var document_prefix = 'document';
    var document_fee_prefix = 'document_fee';
    var lang = $('input[name="selected-lang"]').val();

    // Re-calculation of the price
    function calculateFees() {

        var sub_total = 0;
        var tax_sum = 0;
        var fee_sum = 0;
        $('table.table-formset').each(function () {
            var $table = $(this);
            $table.find('tbody tr[class^="formset_row-"]').each(function () {
                var $tr = $(this);
                var amount = parseInt($tr.find('input[name$="-amount"]').val().replaceAll(',', ''));
                sub_total += amount;
                var tax = parseInt($tr.find('input[name$="-tax"]').val().replaceAll(',', ''));
                tax_sum += tax;
                if ($tr.find('input[name$="-fee"]').length) {
                    var fee = parseInt($tr.find('input[name$="-fee"]').val().replaceAll(',', ''));
                    fee_sum += fee;
                    console.log('[fee sum]', fee_sum);
                }
            });
        });

        $('#id_sub_total').val(sub_total.toLocaleString());
        $('#id_tax').val(tax_sum.toLocaleString());
        if ($('#id_fee').attr('data-type') != 'hall') {
            $('#id_fee').val(fee_sum.toLocaleString());
        }
        calculateTotal();
    }

    // Include/exclude insurance fee
    function calculateTotal() {
        var sub_total = parseInt($('#id_sub_total').val().replaceAll(',', '')) || 0;
        var tax = parseInt($('#id_tax').val().replaceAll(',', '')) || 0;
        var fee = parseInt($('#id_fee').val().replaceAll(',', '')) || 0;
        var total = sub_total + tax + fee;
        $('#id_total').val(total.toLocaleString());
        if ($('#id_billing_amount').length)
            $('#id_billing_amount').val(total.toLocaleString());
    }

    // SetLang
    $('a[data-lang]').click(function (e) {
        e.stopImmediatePropagation();
        e.preventDefault();
        var new_lang = $(e.currentTarget).data('lang');
        if (lang == new_lang) return;

        var url = document.URL.replace(/^(?:\/\/|[^/]+)*\/(ja|en)/, '');

        $.ajax({
            type: 'POST',
            url: '/i18n/setlang/',
            data: {
                language: new_lang,
                next: '/'
            },
            beforeSend: function (request) {
                request.setRequestHeader('X-CSRFToken', csrftoken);
            }
        })
            .done(function () {
                // Remove URL parameters to avoid date format inconsistency b/w English and Japanese
                window.location.href = url.split('?')[0];
            });
    });


    // Delete the selected row and update the formset management
    function deleteRow($table) {
        var $inputTotalForm = $table.find('input[name$="-TOTAL_FORMS"]');
        var numOfItems = parseInt($inputTotalForm.val());
        $inputTotalForm.val(numOfItems - 1);
        var index = 0;
        $table.find('tr[class^="formset_row-"]').each(function (e) {
            $tr = $(this);
            $tr.find('input, select').each(function (e) {
                var el_name = $(this).attr('name');
                el_name = el_name.replace(/\d+/, index);
                $(this).attr('name', el_name);
                $(this).attr('id', 'id_' + el_name);
            });
            index++;
            $tr.find('select.product-type-selectbox').selectBoxIt('destroy');
            $tr.find('select.product-type-selectbox').selectBoxIt({
                autoWidth: false
            });
        });
    }

    // Clicking on delete(-) button in each row inside product/document/documentfee tables
    $('table').on('click', 'a[name="delete_data"]', function (e) {
        e.preventDefault();
        $this = $(this);
        $tr = $this.closest('tr');
        $table = $tr.closest('table');
        $tr.remove();
        deleteRow($table);
        calculateFees();
    });

    // In Sender forms, handles the events when selecting sender
    $('select.select-sender').change(function (e) {
        var id = $(this).val();
        var $self = $(this);
        var $fs = $self.closest('fieldset');

        if (id == "") {
            $fs.find('textarea[name$="_sender_address"]').val(null);
            $fs.find('input[name$="_sender_tel"]').val(null);
            $fs.find('input[name$="_sender_fax"]').val(null);
            return;
        }

        $.ajax({
            type: 'POST',
            url: `/${lang}/master/sender/`,
            data: {
                id: id,
            },
            beforeSend: function (request) {
                request.setRequestHeader('X-CSRFToken', csrftoken);
            },
            success: function (result) {
                $fs.find('textarea[name$="_sender_address"]').val(result.address);
                $fs.find('input[name$="_sender_tel"]').val(result.tel);
                $fs.find('input[name$="_sender_fax"]').val(result.fax);
            }
        });
    });

    // Adding the product to ProductFormSet based table
    $('button[name="add_product_btn"]').click(function (e) {
        // unless product is selected, nothing happens
        var value = $('select.select-product').val();
        if (value == "") {
            $('#modal_product_error').modal('toggle');
            return;
        }
        // reset total number of forms in management form section if there is any cached value
        // after adding selected product name, reset the select2 back to empty option
        resetMangementForm(product_prefix);
        var data = $('select.select-product').select2('data');
        var product = data[0].name;
        $('select.select-product').val(null).trigger('change');
        if ($('table.table-product .odd').length) {
            $('table.table-product .odd').remove();
        }

        // Clone the hiddent formset-row and populate the selected product id/name into the cloned td fields
        var formNum = parseInt($('#id_' + product_prefix + '-TOTAL_FORMS').val());
        var $hiddenTR = $('table.table-product #' + product_prefix + '-formset-row');
        var $newTR = $hiddenTR.clone().removeAttr('style').removeAttr('id').addClass('formset_row-' + product_prefix);
        $('#id_' + product_prefix + '-TOTAL_FORMS').val(formNum + 1);
        $hiddenTR.before($newTR);
        var trRegex = RegExp(`${product_prefix}-xx-`, 'g');
        var html = $newTR.html().replace(trRegex, `${product_prefix}-${formNum}-`);
        $newTR.html(html);
        var productID = `#id_${product_prefix}-${formNum}-product_id`;
        $(productID).val(value);
        var productName = `#id_${product_prefix}-${formNum}-name`;
        $(productName).val(product);

        // after population, make the selectbox inside cloned tr work.
        $newTR.find("select").addClass('product-type-selectbox').selectBoxIt({
            autoWidth: false
        });
    });

    // When clicking "Add Document" button
    $('button[name="add_document_btn"]').click(function (e) {
        // unless any document is selected, nothing happens
        var value = $('select.select-document').val();
        var document = $('select.select-document').children("option:selected").text();
        if (value == "") {
            $('#modal_document_error').modal('toggle');
            return;
        }

        // reset total number of forms in management form section if there is any cached value
        // after adding selected product name, reset the selectbox
        resetMangementForm(document_prefix);
        if ($('table.table-document .odd').length) {
            $('table.table-document .odd').remove();
        }
        $('select.select-document').val("").trigger('change');

        // Clone the hiddent formset-row and populate the selected document id/name into the cloned td fields
        var formNum = parseInt($('#id_' + document_prefix + '-TOTAL_FORMS').val());
        var $hiddenTR = $('table.table-document #' + document_prefix + '-formset-row');
        var $newTR = $hiddenTR.clone().removeAttr('style').removeAttr('id').addClass('formset_row-' + document_prefix);
        $('#id_' + document_prefix + '-TOTAL_FORMS').val(formNum + 1);
        $hiddenTR.before($newTR);
        var trRegex = RegExp(`${document_prefix}-xx-`, 'g');
        var html = $newTR.html().replace(trRegex, `${document_prefix}-${formNum}-`);
        $newTR.html(html);
        var documentID = `#id_${document_prefix}-${formNum}-document_id`;
        $(documentID).val(value);
        var documentName = `#id_${document_prefix}-${formNum}-name`;
        $(documentName).val(document);
        $.ajax({
            type: 'POST',
            url: `/${lang}/contract/check-taxable/`,
            data: {
                id: value,
            },
            beforeSend: function (request) {
                request.setRequestHeader('X-CSRFToken', csrftoken);
            },
            success: function (result) {
                var taxable = result.taxable;
                $(`#id_${document_prefix}-${formNum}-taxable`).val(taxable);
            }
        });
    });

    // When clicking "Add Document Fee" button
    $('button[name="add_document_fee_btn"]').click(function (e) {
        // unless any document is selected, nothing happens
        var value = $('select.select-document-fee').val();
        var document_fee = $('select.select-document-fee').children("option:selected").text();
        if (value == "") {
            $('#modal_document_fee_error').modal('toggle');
            return;
        }
        resetMangementForm(document_fee_prefix);
        if ($('table.table-document_fee .odd').length) {
            $('table.table-document_fee .odd').remove();
        }
        $('select.select-document-fee').val("").trigger('change');
        // Clone the hiddent formset-row and populate the selected document id/name into the cloned td fields
        var formNum = parseInt($('#id_' + document_fee_prefix + '-TOTAL_FORMS').val());
        var $hiddenTR = $('table.table-document_fee #' + document_fee_prefix + '-formset-row');
        var $newTR = $hiddenTR.clone().removeAttr('style').removeAttr('id').addClass('formset_row-' + document_fee_prefix);
        $('#id_' + document_fee_prefix + '-TOTAL_FORMS').val(formNum + 1);
        $hiddenTR.before($newTR);
        var trRegex = RegExp(`${document_fee_prefix}-xx-`, 'g');
        var html = $newTR.html().replace(trRegex, `${document_fee_prefix}-${formNum}-`);
        $newTR.html(html);
        var documentFeeID = `#id_${document_fee_prefix}-${formNum}-document_fee_id`;
        $(documentFeeID).val(value);
        var documentFeeName = `#id_${document_fee_prefix}-${formNum}-name`;
        $(documentFeeName).val(document_fee);

        $.ajax({
            type: 'POST',
            url: `/${lang}/master/document-fee/`,
            data: {
                id: value,
            },
            beforeSend: function (request) {
                request.setRequestHeader('X-CSRFToken', csrftoken);
            },
            success: function (result) {
                var model_price = result.model_price;
                var unit_price = result.unit_price;
                var application_fee = result.application_fee;
                $(`#id_${document_fee_prefix}-${formNum}-model_price`).val(model_price);
                $(`#id_${document_fee_prefix}-${formNum}-unit_price`).val(unit_price);
                $(`#id_${document_fee_prefix}-${formNum}-application_fee`).val(application_fee);
            }
        });
    });

    // Only ASCII charactar in that range allowed
    $('table').on('keypress', 'input[type="number"]', function (e) {
        var ASCIICode = (e.which) ? e.which : e.keyCode;
        if (ASCIICode > 31 && (ASCIICode < 48 || ASCIICode > 57))
            return false;
        return true;
    });


    // Adding change event lister to input field inside table-product
    $('table.table-product').on('input', 'input[type="number"]', function (e) {
        // Calculate quantity * price and set it in id_product-xx-amount td element
        var $self = $(this);
        var $tr = $self.closest('tr');
        var trClassName = $tr.attr('class');
        if (trClassName.startsWith('formset_row-') == false) return;
        var price = parseInt($tr.find('input[name$="-price"]').val().replaceAll(',', ''));
        var quantity = parseInt($tr.find('input[name$="-quantity"]').val().replaceAll(',', ''));
        var amount = parseInt(price) * parseInt(quantity);
        var tax = parseInt(amount * 0.1);
        var fee = 100 * quantity;
        var rounded_price = Math.ceil(price / 10000) * 10000;
        // var rounded_price = price;
        if (rounded_price > 100000) {
            fee = Math.round(200 * quantity * (rounded_price / 100000));
        }
        $tr.find('input[name$="-amount"]').val(amount.toLocaleString());
        $tr.find('input[name$="-tax"]').val(tax.toLocaleString());
        $tr.find('input[name$="-fee"]').val(fee.toLocaleString());
        calculateFees();
    });

    $('table.table-document').on('input', 'input[type="number"]', function (e) {
        var $self = $(this);
        var $tr = $self.closest('tr');
        var trClassName = $tr.attr('class');
        if (trClassName.startsWith('formset_row-') == false) return;
        var price = $tr.find('input[name$="-price"]').val().replaceAll(',', '');
        var quantity = $tr.find('input[name$="-quantity"]').val().replaceAll(',', '');
        var amount = parseInt(price) * parseInt(quantity);
        var taxable = parseInt($tr.find('input[name$="-taxable"]').val().replaceAll(',', ''));
        var tax = parseInt(taxable * amount * 0.1);
        $tr.find('input[name$="-amount"]').val(amount);
        $tr.find('input[name$="-tax"]').val(tax);
        calculateFees();
    });

    $('table.table-document_fee').on('input', 'input[type="number"]', function (e) {
        var $self = $(this);
        var $tr = $self.closest('tr');
        var trClassName = $tr.attr('class');
        if (trClassName.startsWith('formset_row-') == false) return;
        var modelCount = $tr.find('input[name$="-model_count"]').val().replaceAll(',', '');
        var unitCount = $tr.find('input[name$="-unit_count"]').val().replaceAll(',', '');
        var applicationFee = $tr.find('input[name$="-application_fee"]').val().replaceAll(',', '');
        var modelPrice = $tr.find('input[name$="-model_price"]').val().replaceAll(',', '');
        var unitPrice = $tr.find('input[name$="-unit_price"]').val().replaceAll(',', '');
        var amount = parseInt(modelCount) * parseInt(modelPrice) + parseInt(unitCount) * parseInt(unitPrice) + parseInt(applicationFee);
        var tax = parseInt(amount * 0.1);
        $tr.find('input[name$="-amount"]').val(amount);
        $tr.find('input[name$="-tax"]').val(tax);
        calculateFees();
    });

    // when insurance fee value changes...
    $('#id_fee').on('input', function () {
        calculateTotal();
    });

    // when fee_fre checkbox is checked/unchecked...
    $('#fee_editable').change(function () {
        if (this.checked) $('#id_fee').prop('readonly', false); else $('#id_fee').prop('readonly', true);
    });

    // Contract pages submit event
    $('form[name="trader_sales"], form[name="trader_purchases"], form[name="hall_sales"], form[name="hall_purchases"]').on('submit', function (e) {
        var invoiceData = $(this).data('invoice');
        if (invoiceData == false)
            $(this).attr('action', $(location).attr('href'));
        else
            $(this).data('invoice', false);
    });
});
