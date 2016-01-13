function ajustarIframe (iframeWindow) {
  if (iframeWindow.document.height) {
     var iframeElement = document.getElementById(iframeWindow.name);
     iframeElement.style.height = iframeWindow.document.height + 'px';
//     iframeElement.style.width = iframeWindow.document.width + 'px';
  }
  else if (document.all) {
    var iframeElement = document.all[iframeWindow.name];
    if (iframeWindow.document.compatMode &&
        iframeWindow.document.compatMode != 'BackCompat') 
    {
      iframeElement.style.height = iframeWindow.document.documentElement.scrollHeight + 0 + 'px';
//      iframeElement.style.width  = iframeWindow.document.documentElement.scrollWidth + 0 + 'px';
    }
    else {
      iframeElement.style.height = iframeWindow.document.body.scrollHeight + 0 + 'px';
//      iframeElement.style.width = iframeWindow.document.body.scrollWidth + 0 + 'px';
    }
  }
}	
