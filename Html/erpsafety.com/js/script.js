// JavaScript Document

$( document ).ready(function() {


$("a.btn").click(
  function () {
    $(this).addClass('button-hover');
  } 
  
);



/*jQuery function to Fade element */
$( "a.fademe" ).click(function( event ) {
 
    event.preventDefault();
 
    $( this ).hide( "slow" );
 
});


$( "a.alert" ).click(function( event ) {
 
        alert( "Ah lawd jeezus, it's a firrah" );
 
    });
	
	
$(function() {
    $("body").addClass("original");
    $("#btnChange").click(function(e) {
        $("body").toggleClass( "original change", 2000 );
    });
});

    
});