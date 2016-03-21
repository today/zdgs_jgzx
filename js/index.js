/*!
 * index.js v1.0.1
 * (c) 2016 Jin Tian
 * Released under the GPL License.
 */
var xlsx = require("node-xlsx");

var ORDER_DETAIL = null;
var TAX_RATE = 1.17 ;

var init_100 = function(){
  //MSG.put("系统启动。");

  // 充当数据源的文件 的文件名 的关键字
  vm.src_files_flag.push("销售日报表");
  vm.src_files_flag.push("线下订单汇总");
  return true;
}

// 程序运行所必须的库和配置文件
var check_env_110 = function(){

  var envlist = [];
  envlist.push("config.json");
  envlist.push("node_modules/node-xlsx/");
  envlist.push("node_modules/underscore/underscore-min.js");

  return true;
};

var select_file_120 = function(){
  
  if(fs.existsSync(vm.sales_filename)) {
    console.log('销售记录文件存在');
    
  } else {
    console.log('销售记录文件不存在');
    document.getElementById("file_src").click();
  }
  return true;
};

var check_src_130 = function(){
  var run_flag = true;

  // 设置base_dir
  var temp_path = document.getElementById("file_src").value;
  vm.base_dir = path.dirname(temp_path ) + "/" ;
  console.log("base_dir: " + vm.base_dir );

  // 程序运行所必须的数据源
  vm.src_files = find_src_file(vm.base_dir, vm.src_files_flag);
  
  console.log( vm.src_files);

  for( temp_name in vm.src_files){
    if( undefined === vm.src_files[temp_name]){
      ERR_MSG.put( "输入文件不全。缺少：" + temp_name );
      run_flag = false;
    }
  }
  return run_flag;
}

// 装入数据，检查格式
var check_column_140 = function(){
  var run_flag = true;
  
  // 销售日报表  格式检查
  var must_field = [];
  must_field.push("客户名称");
  must_field.push("实际交货数量");
  must_field.push("销售价格");
  must_field.push("物料组描述");
  must_field.push("城市");      // 地市  是否更好
  must_field.push("送货地址");
  must_field.push("物流通知");

  //  线下订单汇总  格式检查
  var must_field2 = [];
  must_field2.push("实际交货数量");
  must_field2.push("销售价格");
  must_field2.push("物料组描述");
  must_field2.push("地市");
  
  var obj = null;
  var order_EOD = vm.src_files["销售日报表"];   //  EOD = End Of Day
  obj = xlsx.parse( vm.base_dir + order_EOD ); // 读入xlsx文件
  //console.log(obj[0]);
  
  MSG.put("开始检查：" + order_EOD );
  // 取出第一个sheet 的第一行。检查标题栏的内容是否正确
  var order_info = obj[0].data;
  var all_title = order_info[0];
  run_flag = check_must_title(all_title, must_field);

  // var obj = null;
  // var order_offline = vm.src_files["线下订单汇总"];   //  EOD = End Of Day
  // obj = xlsx.parse( vm.base_dir + order_offline ); // 读入xlsx文件
  // console.log(order_offline);
  // MSG.put("开始检查：" + order_offline );
  // // 取出第一个sheet 的第一行。检查标题栏的内容是否正确
  // var order_info = obj[0].data;
  // var all_title = order_info[0];
  // run_flag = check_must_title(all_title, must_field);

  // 取出必要的数据
  var index_must_fields = find_title_index_from_array(all_title, must_field);
  console.log(index_must_fields);
  ORDER_DETAIL = select_col_from_array(obj[0].data, index_must_fields);
  //console.log(ORDER_DETAIL);
  
  setTimeout(function() {
    document.getElementById('srcfile_area').style.cssText = "font-size:9px;color:grey;";
    console.log( document.getElementById('srcfile_area').style );
  }, 3*1000);
  return true; //run_flag;
};

// 检查数据是否正确
var check_data_150 = function(){

  // 清洗订单数据
  var index_city = find_title_index(ORDER_DETAIL[0], "城市");
  for(var i=1; i<ORDER_DETAIL.length; i++){
    var order = ORDER_DETAIL[i];
    order[index_city] = get_city(order[index_city]);
  }

  return true;
};

//  计算物料组分项汇总
var calc_160 = function(){

  var all_title = ORDER_DETAIL[0];
  console.log(all_title);
  var city_index = find_title_index(all_title, "城市");
  var prod_index = find_title_index(all_title, "物料组描述");
  var count_index = find_title_index(all_title, "实际交货数量");

  // 获得不重复的地市列表
  var city_list = select_one_col_from_table( ORDER_DETAIL, city_index);
  city_list = _.rest(city_list);  // 去除第0个元素：标题行
  city_list = _.unique(city_list);
  city_list.sort();
  console.log(city_list);
  MSG.put("城市： [" + city_list.join() + "]");

  // 获得不重复的机型列表
  var all_prod_type = select_one_col_from_table( ORDER_DETAIL, prod_index);
  for(i=0;i<all_prod_type.length;i++){
    if( undefined === all_prod_type[i] || "" === all_prod_type[i] ){
      console.log("undefined at " + i );
      ERR_MSG.put("数据错误：发现错误「物料组描述」数据 " + ORDER_DETAIL[i] );
    }
  }
  //console.log(all_prod_type);
  all_prod_type = _.rest(all_prod_type);  // 去除第0个元素：标题行
  all_prod_type = _.unique(all_prod_type);
  all_prod_type.sort();
  console.log(all_prod_type);

  var all_city = [];
  all_city.push("行标签");
  all_city.push("西安城区");
  all_city.push("西安区县");
  all_city.push("咸阳");
  all_city.push("宝鸡");
  all_city.push("渭南");
  all_city.push("铜川");
  all_city.push("延安");
  all_city.push("榆林");
  all_city.push("汉中");
  all_city.push("安康");
  all_city.push("商洛");
  
  function make_line(title_length){
    var line = [];
    for(var i=0; i<title_length; i++){
      line.push(0);
    }
    return line;
  };
  // 构造空白的二维数组
  var result_array = [];
  result_array.push(all_city);
  for(var i=0; i<all_prod_type.length; i++ ){
    var prod = all_prod_type[i];
    var temp = make_line(all_city.length);
    temp[0] = prod;
    result_array.push(temp);
  }

  // 填充数据
  var result_title = result_array[0];
  for(var i=1;i<ORDER_DETAIL.length;i++){
    var order = ORDER_DETAIL[i];
    var city = order[city_index];
    var prod = order[prod_index];
    var count = order[count_index];

    for(var j=1; j<result_array.length; j++){
      var temp = result_array[j];
      var prod2 = temp[0];
      
      if( prod === prod2 ){
        for(var k=1; k<result_title.length; k++ ){
          var city2 = result_title[k];
          if( city === city2 ){
            temp[k] += count;
            console.log(temp[k]);
          }
        }
        break;
      }
    }
  }
  
                      
  // 把计算结果存入文件。
  var buffer = xlsx.build([{name: "汇总日报", data: result_array }] );
  fs.writeFileSync( vm.base_dir + "中间文件_第一步.xlsx", buffer);

  return true;
}


var get_city = function( src_city ){
    var dest_city = "出错";
    if( "" === src_city ) city = "";
    else if ( src_city.indexOf("西安城区") > -1 ) city = "西安城区";
    else if ( src_city.indexOf("西安市") > -1 ) city = "西安城区";
    else if ( src_city.indexOf("市场部大客户") > -1 ) city = "西安城区";
    else if ( src_city.indexOf("西安区县") > -1 ) city = "西安区县";
    else if ( src_city.indexOf("西安郊县") > -1 ) city = "西安区县";
    else if ( src_city.indexOf("咸阳") > -1 ) city = "咸阳";
    else if ( src_city.indexOf("宝鸡") > -1 ) city = "宝鸡";
    else if ( src_city.indexOf("渭南") > -1 ) city = "渭南";
    else if ( src_city.indexOf("铜川") > -1 ) city = "铜川";
    else if ( src_city.indexOf("延安") > -1 ) city = "延安";
    else if ( src_city.indexOf("榆林") > -1 ) city = "榆林";
    else if ( src_city.indexOf("汉中") > -1 ) city = "汉中";
    else if ( src_city.indexOf("安康") > -1 ) city = "安康";
    else if ( src_city.indexOf("商洛") > -1 ) city = "商洛";
    else{
      city = "出错";
      console.log("「地市」出错。 #" + src_city + "#");
      ERR_MSG.put("「地市」出错。 #" + src_city + "#");
    } 
    return city;
  };


var check_must_title = function(target, keywords ){
  var run_flag = true;
  for(var i=0; i<keywords.length; i++ ){
    var kw = keywords[i];
    if( -1 === _.indexOf(target, kw)){
      // 显示出错提示。
      ERR_MSG.put("数据出错：。无法找到「"+ kw +"」列，请检查数据的第一行。" );
      console.log("数据出错：。无法找到「"+ kw +"」列，请检查数据的第一行。");
      console.log(target);
      run_flag = false;
    }
  }
  return run_flag;
}









