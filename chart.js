$(document).ready(function(){
    $("#btn_show1").click(function(){
      show_string = ''
      $(":checked").each(function(){
        show_string = show_string + $(this).attr("id") + '<br/>'
      });
      //alert(show_string)
      $("#result1").html(show_string);
    });
    $("#btn_show2").click(function(){
      show_string = ''
      $(":checked").each(function(){
        show_string = show_string + $(this).attr("id") + '<br/>'
      });
      //alert(show_string)
      $("#result2").html(show_string);
    });
  });