/*!
 * lib.js v1.0.1
 * (c) 2015 Jin Tian
 * Released under the GPL License.
 */
var _ = require('underscore');
var fs = require('fs');
var path = require('path');

var Msg = {
  createNew: function(){
    var msg = {};
    msg.msg_list = [];
    msg.msg_count = 0;

    msg.put = function(a_msg){ 
      //msg.msg_list.push("log " + msg.msg_list.length + ": " + a_msg); 
      msg.msg_count++;
      msg.msg_list.splice(0,0,"\nlog " + msg.msg_count + ": " + a_msg)
      //msg.msg_list.push( a_msg ); 
      
    };

    msg.download = function(){
      var evt = document.createEvent("HTMLEvents");
      evt.initEvent("click", false, false);//initEvent 不加后两个参数在FF下会报错

      var blob = new Blob(msg.msg_list);

      var aLink = document.createElement('a');
      aLink.download = "错误数据记录.txt";
      aLink.href = URL.createObjectURL(blob, { "type" : "text/xml" });
      aLink.dispatchEvent(evt);
    }

    return msg;
  }
};

function check_index(a_array){
  var flag = true;
  for( temp in a_array ){
    //console.log(temp);
    if( a_array[temp] === -1 ){
      flag = false;
    }
  }
  return flag;
}

// 根据关键词，寻找文件名匹配的目标文件。
function find_src_file( base_dir, files_flag ){
  var dest_files = [];

  if(fs.existsSync(base_dir)) {
    console.log('base_dir 存在');

    var all_file = fs.readdirSync(base_dir);
    //console.log(files_flag);

    for(var i=0; i<files_flag.length; i++){
      var key=files_flag[i];
      var file_name = undefined;

      for(var j=0; j<all_file.length; j++){
        if( all_file[j].indexOf(key) > -1 ){
          var file_name = all_file[j];
        }
      }
      dest_files[key] = file_name;
    }
  } else {
    console.log('base_dir 不存在');
    return null;
  }
  //console.log(dest_files);
  return dest_files;
}

/*
 从二维数组中删除指定的列 
*/
function del_col_from_array(a_array, col_indexs){
  //console.log(a_array);
  var src_col_index = _.range(a_array[0].length);
  var select_index = _.difference(src_col_index, col_indexs);
  return select_col_from_array(a_array, select_index);
}

/*
 从二维数组中取出指定的列 
*/
function select_col_from_array(a_array, col_indexs){
  var dest = [];
  var temp = [];
  for(var i=0; i<a_array.length; i++){
    temp = [];
    for(var j=0; j<col_indexs.length; j++){
      temp.push(a_array[i][col_indexs[j]]);
    }
    dest.push(temp);
  }
  return dest;
}

/*
 从二维数组中取出指定的一个列，并做成一维数组返回
*/
function select_one_col_from_table(a_array, col_index ){
  var dest = [];
  var temp = null;
  for(var i=0; i<a_array.length; i++){
    temp = null;
    temp = a_array[i][col_index];
    
    dest.push(temp);
  }
  return dest;
}

/*
 给二维数组加一列   在尾部
*/
function add_col_for_table(a_array, default_val){
  var dest = [];
  if(default_val === null ){
    default_val = "";
  }
  for(var i=0; i<a_array.length; i++){
    var temp = a_array[i];
    temp.push(default_val);
    dest.push(temp);
  }
  return dest;
}

function trim_array_element( a_array ){
  for( var i=0; i<a_array.length; i++){
    var temp = a_array[i];
    if( temp ){
      if( (typeof temp) == "string" ){
        a_array[i] = temp.trim();
      }else{
        a_array[i] = temp.toString().trim();
      }
    }
  }
}

// 找出单个index值
var find_title_index = function( title_array, t_name){
  var a_index = _.indexOf( title_array, t_name );
  return a_index;
};
// 找出一组个index值
var find_title_index_from_array = function( title_array, name_array ){
  var index_array = [];
  for( var i=0; i<name_array.length; i++ ){
    index_array.push(find_title_index(title_array, name_array[i]));
  }
  return index_array;
};


function isblank(strA){
  if(strA){
    if( "string" === typeof(strA) ){
      if( "" === strA.trim()){
        return true;
      }else{
        return false;
      }
    }else{
      return false;
    }
  }else{
    return true;
  }
}