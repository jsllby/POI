/**
 * Created by fangzhiyang on 2017/4/25.
 */
jQuery(document).ready(function() {
	loadCount();
	loadInfo();

	jQuery('#export1').click(function() {
    	onExport(1);
    });
    
	jQuery('#export2').click(function() {
    	onExport(2);
    });

	jQuery('#export3').click(function() {
    	onExport(3);
    });
	
	jQuery('#export4').click(function() {
    	onExport(4);
    });
});

// 经办人填写
function agent(panel) {
    var name = jQuery("#name" + panel).val();
    var tel = jQuery("#tel" + panel).val();
    jQuery("#agent" + panel).val(name + tel);
}

// 提交
function onExport(panel) {
    switch (panel){
        case 1:
            if (jQuery("#itemcode1").val() == "") {
                alert("项目代码必须填写哦！");
                return;
            }
            if (jQuery("#realAmount1").val() == "") {
                alert("报账金额必须填写哦！");
                return;
            }
            if (eval(jQuery("#realAmount1").val()) > 99999.99) {
                alert("报账金额不能大于99999.99，请填写大额报账单！");
                return;
            }
            break;
        case 2:
            if (jQuery("#itemcode2").val() == "") {
                alert("项目代码必须填写哦！");
                return;
            }
            if (jQuery("#realAmount2").val() == "") {
                alert("报账金额必须填写哦！");
                return;
            }
            if (eval(jQuery("#realAmount2").val()) > 99999.99) {
                alert("报账金额不能大于99999.99，请填写大额报账单！");
                return;
            }
            break;
        case 3:
            if (jQuery("#itemcode3").val() == "") {
                alert("项目代码必须填写哦！");
                return;
            }
            if (jQuery("#realAmount3").val() == "") {
                alert("报账金额必须填写哦！");
                return;
            }
            if (eval(jQuery("#realAmount3").val()) > 999999999.99) {
                alert("报账金额不能大于999999999.99!");
                return;
            }
            break;
        case 4:
            if (jQuery("#itemcode4").val() == "") {
                alert("项目代码必须填写哦！");
                return;
            }
            if (jQuery("#realAmount4").val() == "") {
                alert("报账金额必须填写哦！");
                return;
            }
            if (eval(jQuery("#realAmount4").val()) > 999999999.99) {
                alert("报账金额不能大于999999999.99!");
                return;
            }
            break;
        case 5:
            if (jQuery("#itemcode5").val() == "") {
                alert("项目代码必须填写哦！");
                return;
            }
        case 6:
            if (jQuery("#itemcode6").val() == "") {
                alert("项目代码必须填写哦！");
                return;
            }
    }
    
    // 保存localStorage，保存用户输入的数据
    saveInfo();
    
    jQuery("#export_form" + panel).submit();
}

function loadCount() {
    var callback = function(result)
    {
        jQuery("#count").text(result.count);
    }
    var operation = new Operation("公用.获取统计数");
    operation.name = "scufd";
	operation.execute(callback);
}

// 加载本地存储数据
function loadInfo() {
    var elements = jQuery("form :input");
    jQuery.each(elements, function(i, element){
    	if (jQuery(element).val() == "")
    		jQuery(element).val(localStorage.getItem(element.name));
    });
}

// 保存本地数据
function saveInfo() {
    var fields = jQuery(":input").serializeArray();
    jQuery.each(fields, function(i, field){
    	if (field.name != "jobnum" || field.value != "")
    		localStorage.setItem(field.name, field.value);
    });
}

var index = 1;
// 添加行程
function addTrip() {
	var html = [];
	html.push('<div class="control-group">');
	html.push('  <label class="control-label">行程' + index++ + '</label>');
	html.push('  <div class="controls">');
	html.push('    <input name="test1" type="text" value="" class="input-xlarge"/>');
	html.push('    <input name="test2" type="text" value="" class="input-xlarge"/>');
	html.push('    <input name="test3" type="text" value="" class="input-xlarge"/>');
	html.push('    <input type="button" value="删除" class="btn btn-success"/>');
	html.push('  </div>');
	html.push('</div>');
	
	var trip = jQuery(html.join(''));
	trip.find(':input[type=button]').click(function() {
		trip.empty();
	});
	jQuery('#container').append(trip);
}

function deleteTrip() {
	jQuery('#container').empty();
}