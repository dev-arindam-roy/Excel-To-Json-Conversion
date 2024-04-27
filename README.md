# Excel To JSON Conversion

## An application on excel to json conversion

```js
$(document).ready(function () {

    window.addEventListener('load', () => {
        registerSW();
    });

    // Register the Service Worker
    async function registerSW() {
        if ('serviceWorker' in navigator) {
            try {
                await navigator.serviceWorker.register('assets/pwa/serviceworker.js');
            }
            catch (e) {
                console.log('SW registration failed');
            }
        }
    }

    $('.onex-select2').select2({
        width: '100%',
        placeholder: 'Select an option',
        allowClear: false
    });

    $('body').on('select2:select', '.onex-select2', function (e) { 
        if($(this).val() != '') {
            $('#' + $(this).attr('id') + '-error').hide();
            $(this).next('span.select2-container').removeClass('select2-custom-error');
            $(this).parent().find('.onex-form-label').removeClass('onex-error-label');
        }
    });

    $('body').on('select2:selecting', '.onex-select2', function (e) { 
        if($(this).val() != '') {
            $('#' + $(this).attr('id') + '-error').hide();
            $(this).next('span.select2-container').removeClass('select2-custom-error');
            $(this).parent().find('.onex-form-label').removeClass('onex-error-label');
        }
    });

    /** Copy the raw json */
    $('#copyJsonBtn').on('click', async function() {
        let copyJson = document.getElementById('convertedJson');
        await window.navigator.clipboard.writeText(copyJson.value);
        copyJson.select();
        displayToast();
    });

    /** SweetAlert2 loading */
    const displayLoading = (timer = 3000, title = 'Please Wait...', text = "System Processing Your Request") => {
        Swal.fire({
            title: title,
            text: text,
            allowEscapeKey: false,
            allowOutsideClick: false,
            timer: timer,
            didOpen: () => {
                Swal.showLoading()
            }
        });
    }

    /** SweetAlert2 like toast */
    const displayToast = () => {
        Swal.fire({
            position: 'top-end',
            icon: 'success',
            title: 'Its copied!',
            showConfirmButton: false,
            timer: 1000
        });
    }

    $("#frmx").validate({
        errorClass: 'onex-error',
        errorElement: 'div',
        rules: {
            excel_file: {
                required: true,
                extension: 'xls|xlsx'
            }
        },
        messages: {
            excel_file: {
                required: 'Please upload excel file',
                extension: 'Only accept .xls and .xlsx file',
                accept: 'Only accept .xls and .xlsx file'
            }
        },
        errorPlacement: function (error, element) {
            if(element.hasClass('onex-select2')) {
                error.insertAfter(element.parent().find('span.select2-container'));
            } else {
                error.insertAfter(element);
            }
        },
        highlight: function (element) {
            $(element).removeClass('is-valid').addClass('is-invalid');
            $(element).parent().find('.onex-form-label').addClass('onex-error-label');
            if(element.type == 'select-one') {
                $(element).next('span.select2-container').addClass('select2-custom-error');
            }
        },
        unhighlight: function (element) {
            $(element).removeClass('is-invalid').addClass('is-valid');
            $(element).parent().find('.onex-form-label').removeClass('onex-error-label');
        },
        submitHandler: function (form) {
            uploadExcelFile();
            return false;
        }
    });

    $("#frmx2").validate({
        errorClass: 'onex-error',
        errorElement: 'div',
        rules: {
            excel_sheet: {
                required: true
            }
        },
        messages: {
            excel_sheet: {
                required: 'Please select a sheet'
            }
        },
        errorPlacement: function (error, element) {
            if(element.hasClass('onex-select2')) {
                error.insertAfter(element.parent().find('span.select2-container'));
            } else {
                error.insertAfter(element);
            }
        },
        highlight: function (element) {
            $(element).removeClass('is-valid').addClass('is-invalid');
            $(element).parent().find('.onex-form-label').addClass('onex-error-label');
            if(element.type == 'select-one') {
                $(element).next('span.select2-container').addClass('select2-custom-error');
            }
        },
        unhighlight: function (element) {
            $(element).removeClass('is-invalid').addClass('is-valid');
            $(element).parent().find('.onex-form-label').removeClass('onex-error-label');
        }
    });

    function actionButtons(jsonObj) {
        if(jsonObj.length) {
            $('#downloadJsonBtn').removeClass('disabled');
            $('#copyJsonBtn').removeClass('disabled');
        } else {
            $('#downloadJsonBtn').addClass('disabled');
            $('#copyJsonBtn').addClass('disabled');
        }
    }

    let excelFile = document.querySelector('#excelFileUpload');
    let excelSheets = document.querySelector('#excelFileSheets');
    let excel;

    function uploadExcelFile() {
        displayLoading();
        excelFile.files[0].arrayBuffer().then((buffer) => {
            excel = XLSX.read(buffer);
            if((excel.SheetNames) && excel.SheetNames.length) {
                let sheetDropDown = excel.SheetNames.forEach((item, index) => {
                    excelSheets.innerHTML += `<option value="${item}">Sheet: ${item}</option>`;
                });
                $('#excelFileSheets').val(excel.SheetNames[0]).trigger('change');
                $('#excelFileSheets').trigger({type: 'select2:select'});
                let allSheets = excel.Sheets;
                let selectedSheet = allSheets[$('#excelFileSheets').val()];
                let jsonObj = XLSX.utils.sheet_to_json(selectedSheet);
                actionButtons(jsonObj);
                $('#convertedJson').val(JSON.stringify(jsonObj, null, 4));
                $('#json-renderer').jsonViewer(jsonObj);
                $('.second-phase').show();
            }
        });
    }

    $('#excelFileSheets').on('change', function() {
        if($(this).valid() && $('#frmx2').valid()) {
            displayLoading(1000);
            let allSheets = excel.Sheets;
            let selectedSheet = allSheets[$(this).val()];
            let jsonObj = XLSX.utils.sheet_to_json(selectedSheet);
            actionButtons(jsonObj);
            $('#convertedJson').val(JSON.stringify(jsonObj, null, 4));
            $('#json-renderer').jsonViewer(jsonObj);
        }
    });

    $('#downloadJsonBtn').on('click', function() {
        const anchor = document.createElement('a');
        anchor.href = 'data:application/json;charset=utf-8,' + encodeURIComponent($('#convertedJson').val());
        anchor.download = $('#excelFileSheets').val();
        document.body.appendChild(anchor);
        anchor.click();
        document.body.removeChild(anchor);
    });
});
```