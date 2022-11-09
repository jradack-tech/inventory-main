/* ------------------------------------------------------------------------------
*
*  # Select2 selects
*
*  Demo JS code for form_select2.html page
*
* ---------------------------------------------------------------------------- */

document.addEventListener('DOMContentLoaded', function () {

    // Select customer select2 initialization and formatting
    function formatCustomer(customer) {
        if (customer.loading) return customer.text;

        var markup = "<div class='select2-result-customer clearfix'>" +
            "<div class='select2-result-customer__name'><b>" + customer.name + "</b></div>" +
            "<div class='select2-result-customer__frigana'>【" + customer.frigana + "】</div>" +
            "<div class='select2-result-customer__frigana'>【" + customer.address + "】</div>";
        if (customer.tel) {
            markup += "<div class='select2-result-customer__tel'>TEL : <b>" + customer.tel + "</b></div>";
        }
        if (customer.fax) {
            markup += "<div class='select2-result-customer__fax'>FAX : <b>" + customer.fax + "</b></div>";
        }
        markup += "</div>";
        return markup;
    }

    function formatCustomerSelection(customer) {
        return customer.name || customer.text;
    }

    $(".select-customer").select2({
        ajax: {
            url: "/master/search-customer/",
            dataType: 'json',
            delay: 250,
            data: function (params) {
                return {
                    q: params.term,
                    page: params.page
                };
            },
            processResults: function (data, params) {
                params.page = params.page || 1;
                return {
                    results: data.customers,
                    pagination: {
                        more: (params.page * 30) < data.total_count
                    }
                };
            },
            cache: true
        },
        escapeMarkup: function (markup) { return markup; },
        // allowClear: true,
        minimumInputLength: 1,
        templateResult: formatCustomer,
        templateSelection: formatCustomerSelection
    });

    // customer-search select2 changed event
    $('.select-customer').on('select2:select', function (e) {
        console.log('clicked--oks asdf');
        var customer = e.params.data;
        var frigana = customer.frigana;
        var postal_code = customer.postal_code;
        var address = customer.address;
        var tel = customer.tel;
        var fax = customer.fax;


        if (frigana) {
            // updates the values only when admin selects the customer manually
            $('input[name="frigana"]').val(frigana);
            $('input[name="postal_code"]').val(postal_code);
            $('input[name="address"]').val(address);
            $('input[name="tel"]').val(tel);
            $('input[name="fax"]').val(fax);

            // Auto fill product_sender, document_sender
            // $(".select-product-sender").next().find('span.select2-selection__rendered').text(customer.name)
            // $('textarea[name="product_sender_address"]').val(address);
            // $('input[name="product_sender_tel"]').val(tel);
            // $('input[name="product_sender_fax"]').val(fax);

            // $(".select-document-sender").next().find('span.select2-selection__rendered').text(customer.name)
            // $('textarea[name="document_sender_address"]').val(address);
            // $('input[name="document_sender_tel"]').val(tel);
            // $('input[name="document_sender_fax"]').val(fax);

            $('span.buyer-postal-code').text(postal_code);
            $('span.buyer-address').text(address);
            $('h4.buyer-company').text(customer.name);
            $('span.buyer-tel').text(tel);
            $('span.buyer-fax').text(fax);
        }

    });

    $('.v-sender-clear').on('click', function (e) {
        e.preventDefault();
        $('.v-select-product-sender > .select-product-sender').html('<option value="">商品発送先を選択してください...</option>');
        $('textarea[name="product_sender_address"]').val('');
        $('input[name="product_sender_postal_code"]').val('');
        $('input[name="product_sender_tel"]').val('');
        $('input[name="product_sender_fax"]').val('');
    });

    $('.v-document-clear').on('click', function (e) {
        e.preventDefault();
        $('.v-select-document-sender > .select-document-sender').html('<option value="">書類発送先を選択してください...</option>');
        $('textarea[name="document_sender_address"]').val('');
        $('input[name="document_sender_postal_code"]').val('');
        $('input[name="document_sender_tel"]').val('');
        $('input[name="document_sender_fax"]').val('');
    });

    $("#id_shipping_methodSelectBoxItText").on('change', function (e) {
        console.log('[changed]');
    });

    // $('.v-contract-clear').on('click', function(e) {
    //     console.log('[clearing]');
    //     $('.v-select-customer > .select-customer').text('');
    //     $('input[name="frigana"]').val('');
    //     $('input[name="postal_code"]').val('');
    //     $('input[name="address"]').val('');
    //     $('input[name="tel"]').val('');
    //     $('input[name="fax"]').val('');
    //     $('span.buyer-postal-code').text('');
    //     $('span.buyer-address').text('');
    //     $('h4.buyer-company').text(''.name);
    //     $('span.buyer-tel').text('');
    //     $('span.buyer-fax').text('');
    // });
    // End of select customer select2 initialization and formatting




    // Select hall select2 initialization and formatting
    function formatHallCustomer(hall) {
        if (hall.loading) return hall.text;

        var markup = "<div class='select2-result-hall clearfix'>" +
            "<div class='select2-result-hall__customer_name'><b>" + hall.customer_name + "</b></div>" +
            "<div class='select2-result-hall__customer_frigana'>【" + hall.customer_frigana + "】</div>" +
            "<div class='select2-result-hall__name'><b>&nbsp;&nbsp;" + hall.name + "</b></div>" +
            "<div class='select2-result-hall__frigana'>【" + hall.frigana + "】</div>";
        if (hall.tel) {
            markup += "<div class='select2-result-hallr__tel'>TEL : <b>" + hall.tel + "</b></div>";
        }
        if (hall.fax) {
            markup += "<div class='select2-result-hall__fax'>FAX : <b>" + hall.fax + "</b></div>";
        }
        markup += "</div>";
        return markup;
    }

    function formatHallCustomerSelection(hall) {
        return hall.customer_name || hall.text;
    }

    $(".select-hall-customer").select2({
        ajax: {
            url: "/master/search-hall/",
            dataType: 'json',
            delay: 250,
            data: function (params) {
                return {
                    q: params.term,
                    page: params.page
                };
            },
            processResults: function (data, params) {
                params.page = params.page || 1;
                return {
                    results: data.halls,
                    pagination: {
                        more: (params.page * 30) < data.total_count
                    }
                };
            },
            cache: true
        },
        escapeMarkup: function (markup) { return markup; },
        // allowClear: true,
        minimumInputLength: 1,
        templateResult: formatHallCustomer,
        templateSelection: formatHallCustomerSelection
    });

    // customer-search select2 changed event
    $('.select-hall-customer').on('select2:select', function (e) {
        console.log('[hall clicked]');
        var hall = e.params.data;
        var name = hall.name;
        var address = hall.address;
        var tel = hall.tel;

        // if (customer_name) $('input[name="customer_name"]').val(customer_name);
        if (name) $('input[name="hall_name"]').val(name);
        if (address) $('input[name="hall_address"]').val(address);
        if (tel) $('input[name="hall_tel"]').val(tel);
    });

    // Select hall select2 initialization and formatting
    function formatHall(hall) {
        if (hall.loading) return hall.text;

        var markup = "<div class='select2-result-hall clearfix'>" +
            "<div class='select2-result-hall__name'><b>" + hall.name + "</b></div>" +
            "<div class='select2-result-hall__frigana'>【" + hall.frigana + "】</div>";
        if (hall.tel) {
            markup += "<div class='select2-result-hallr__tel'>TEL : <b>" + hall.tel + "</b></div>";
        }
        if (hall.fax) {
            markup += "<div class='select2-result-hall__fax'>FAX : <b>" + hall.fax + "</b></div>";
        }
        markup += "</div>";
        return markup;
    }

    function formatHallSelection(hall) {
        return hall.name || hall.text;
    }

    $(".select-hall").select2({
        ajax: {
            url: "/master/search-hall/",
            dataType: 'json',
            delay: 250,
            data: function (params) {
                return {
                    q: params.term,
                    page: params.page
                };
            },
            processResults: function (data, params) {
                params.page = params.page || 1;
                return {
                    results: data.halls,
                    pagination: {
                        more: (params.page * 30) < data.total_count
                    }
                };
            },
            cache: true
        },
        escapeMarkup: function (markup) { return markup; },
        // allowClear: true,
        minimumInputLength: 1,
        templateResult: formatHall,
        templateSelection: formatHallSelection
    });
    // End of select hall select2 initialization and formatting

    // hall-search select2 changed event
    $('.select-hall').on('select2:select', function (e) {
        var hall = e.params.data;
        var address = hall.address;
        var tel = hall.tel;
        if (address) $('input[name="hall_address"]').val(address);
        if (tel) $('input[name="hall_tel"]').val(tel);
    });


    // Select product select2 initialization and formatting
    function formatProduct(product) {
        if (product.loading) return product.text;

        var markup = "<div class='select2-result-product clearfix'>" +
            "<div class='select2-result-product__name'><b>" + product.name + "</b></div></div>";
        return markup;
    }

    function formatProductSelection(product) {
        return product.name || product.text;
    }

    $('.select-product').select2({
        ajax: {
            url: "/master/search-product/",
            dataType: 'json',
            delay: 250,
            data: function (params) {
                return {
                    q: params.term,
                    page: params.page
                };
            },
            processResults: function (data, params) {
                params.page = params.page || 1;
                return {
                    results: data.products,
                    pagination: {
                        more: (params.page * 30) < data.total_count
                    }
                };
            },
            cache: true
        },
        escapeMarkup: function (markup) { return markup; },
        // allowClear: true,
        minimumInputLength: 1,
        templateResult: formatProduct,
        templateSelection: formatProductSelection
    });
    // End of select hall select2 initialization and formatting

    // Select sender select2 initialization and formatting
    function formatSender(sender) {
        if (sender.loading) return sender.text;

        var markup = "<div class='select2-result-sender clearfix'>" +
            "<div class='select2-result-sender__name'><b>" + sender.name + "</b></div>" +
            "<div class='select2-result-customer__frigana'>【" + sender.address + "】</div></div>";
        return markup;
    }

    function formatSenderSelection(sender) {
        return sender.name || sender.text;
    }

    $(".select-sender").select2({
        ajax: {
            url: "/master/search-sender/",
            dataType: 'json',
            delay: 250,
            data: function (params) {
                return {
                    q: params.term,
                    page: params.page
                };
            },
            processResults: function (data, params) {
                params.page = params.page || 1;
                return {
                    results: data.senders,
                    pagination: {
                        more: (params.page * 30) < data.total_count
                    }
                };
            },
            cache: true
        },
        escapeMarkup: function (markup) { return markup; },
        // allowClear: true,
        minimumInputLength: 1,
        templateResult: formatSender,
        templateSelection: formatSenderSelection
    });

    // sender search select2 changed event
    $('.select-sender').on('select2:select', function (e) {
        var sender = e.params.data;
        var address = sender.address;
        var postal_code = sender.postal_code;
        var tel = sender.tel;
        var fax = sender.fax;

        console.log(sender);

        var $fieldset = $(this).closest('fieldset');
        if (address) $fieldset.find('textarea[name$="_sender_address"]').val(address);
        if (postal_code) $fieldset.find('input[name$="_sender_postal_code"]').val(postal_code);
        if (tel) $fieldset.find('input[name$="_sender_tel"]').val(tel);
        if (fax) $fieldset.find('input[name="_sender_fax"]').val(fax);

    });
    // End of select sender select2 initialization and formatting

});
