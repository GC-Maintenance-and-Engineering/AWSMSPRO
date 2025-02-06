/*
 * Form Validation
 */
$(function () {

    $('select[required]').css({
        position: 'absolute',
        display: 'inline',
        height: 0,
        padding: 0,
        border: '1px solid rgba(255,255,255,0)',
        width: 0
    });

    $("#formValidate").validate({
        rules: {
            ntitle: {
                required: true,
                maxlength: 100
            },
            nreqno: {
                required: true,
                maxlength: 50
            },
            nnumber: {
                required: true,
                number: true
            },
            nfile: {
                required: true
            },
            nenddate: {
                required: true,
                date: true
            },
            nemp_id: {
                required: true,
                number: true,
                minlength:10
            },
            nadds: {
                maxlength: 200
            },
            nclentname: {
                required: true,
                maxlength: 200
            },
            npersonname: {
                required: true,
                maxlength: 100
            },
            ntel: {
                phone:true
            },
            nmobile: {
                phone: true
            },
            nrem: {
                maxlength: 250
            },
            tnc_select: "required",
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });
    $("#formValidateSucc").validate({
        rules: {
            nemp_id: {
                required: true,
                number: true,
                minlength: 8
            },
            sel_cand: {
                required: true,
            },
            tnc_select: "required",
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });
    $("#formValidateUploadFile").validate({
        rules: {
            nfile: {
                required: true,
            }
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });
    $("#formValidatePps").validate({
        rules: {
            start_date: {
                required: true,
            },
            end_date: {
                required: true,
            },
            working_date: {
                required: true,
            },
            proposal_date: {
                required: true,
            },
            scope_service: {
                required: true,
            },
            act_salary: {
                required: true,
            },
            period: {
                required: true,
            },
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });
    $("#formValidateCand").validate({
        rules: {
            f_name: {
                required: true,
                maxlength: 65
            },
            l_name: {
                required: true,
                maxlength: 65
            },
            sel_gender: {
                required: true,
            },
            tel: {
                required: true,
                phone: true
            },
            tel_s: {
                phone: true
            },
            mail: {
                maxlength: 50
            },
            id_card: {
                maxlength: 13
            },
            remark: {
                maxlength: 512
            }
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });
    $("#formValidateEdu").validate({
        rules: {
            degree: {
                required: true,
            },
            inst: {
                required: true,
                maxlength: 175
            },
            faculty: {
                required: true,
                maxlength: 75
            },
            major: {
                required: true,
                maxlength: 75
            }
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });

    $("#formValidateApply").validate({
        rules: {
            app_posi: {
                required: true,
                maxlength: 100
            },
            app_date: {
                required: true,
            },
            cv_file: {
                required: true,
            },
            curr_work: {
                maxlength: 75
            },
            keyword: {
                maxlength: 512
            },
            app_remark: {
                maxlength: 256
            }
        },
        //For custom messages
        messages: {
            uname: {
                required: "Enter a username",
                minlength: "Enter at least 5 characters"
            },
            curl: "Enter your website",
        },
        errorElement: 'div',
        errorPlacement: function (error, element) {
            var placement = $(element).data('error');
            if (placement) {
                $(placement).append(error)
            } else {
                error.insertAfter(element);
            }
        }
    });
   
    $.validator.addMethod("dateTime", function (value, element) {
        //var stamp = value.split(" ");
        //var validDate = !/Invalid|NaN/.test("dd M yy",new Date(stamp[0]).toString());
        // var validTime = /^(([0-1]?[0-9])|([2][0-3])):([0-5]?[0-9])(:([0-5]?[0-9]))?$/i.test(stamp[1]);
        //return this.optional(element) || parseDate(value, "dd/MM/yyyy") !== null;
        var check = false,
            re = /^\d{1,2}\/\d{1,2}\/\d{4}$/,
            adata, gg, mm, aaaa, xdata;
        if (re.test(value)) {
            adata = value.split("/");
            gg = parseInt(adata[0], 10);
            mm = parseInt(adata[1], 10);
            aaaa = parseInt(adata[2], 10);
            xdata = new Date(aaaa, mm - 1, gg, 12, 0, 0, 0);
            if ((xdata.getUTCFullYear() === aaaa) && (xdata.getUTCMonth() === mm - 1) && (xdata.getUTCDate() === gg)) {
                check = true;
            } else {
                check = false;
            }
        } else {
            check = false;
        }
        return this.optional(element) || check;
    }, "Please enter a valid date.");

    $.validator.addMethod("date", function (value, element) {
        var dtValue = value;
        var dtRegex = new RegExp("^([0]?[1-9]|[1-2]\\d|3[0-1])-(JAN|FEB|MAR|APR|MAY|JUN|JULY|AUG|SEP|OCT|NOV|DEC)-[1-2]\\d{3}$", 'i');
        return dtRegex.test(dtValue);
    }, "Please enter a valid date.");

    jQuery.validator.addMethod("phone", function (value, element) {
        //!/^\d{8}$|^\d{10}$/.test(value) ? false : true;
        return this.optional(element) || /^((\+|00(\s|\s?\-\s?)?)31(\s|\s?\-\s?)?(\(0\)[\-\s]?)?|0)[1-9]((\s|\s?\-\s?)?[0-9]){8}$/.test(value);
    }, "Please specify a valid phone number.");
});