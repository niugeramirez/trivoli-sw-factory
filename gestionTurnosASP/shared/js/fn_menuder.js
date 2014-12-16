var objmenuderecho;

function showmenuder(){
  var menuder = objmenuderecho;
  var rightedge=document.body.clientWidth-event.clientX
  var bottomedge=document.body.clientHeight-event.clientY

  if (rightedge<menuder.offsetWidth){
	 if ((document.body.scrollLeft+event.clientX-menuder.offsetWidth) < document.body.scrollLeft){
        menuder.style.left=document.body.scrollLeft;
	 }else{
        menuder.style.left=document.body.scrollLeft+event.clientX-menuder.offsetWidth;
	 }
  }else{
     menuder.style.left=document.body.scrollLeft+event.clientX;
  }

  if (bottomedge<menuder.offsetHeight){
	 if ((document.body.scrollTop+event.clientY-menuder.offsetHeight) < document.body.scrollTop){
        menuder.style.top=document.body.scrollTop
	 }else{
        menuder.style.top=document.body.scrollTop+event.clientY-menuder.offsetHeight;	   
	 }
  }else{
     menuder.style.top=document.body.scrollTop+event.clientY;
  }	 
  
  menuder.style.visibility="visible"

  return false
}

function hidemenuder(){
  var menuder = objmenuderecho;

  menuder.style.visibility="hidden";
}

function inicializarMenuDerecho(obj){
  objmenuderecho = obj;
  document.oncontextmenu=showmenuder;
  document.body.onclick=hidemenuder;
}

