function swap(targetId){
  
  if (document.getElementById)
        {
        target = document.getElementById(targetId);
        
            if (target.style.display == "none")
                {
                target.style.display = "";
                } 
            else 
                {
                target.style.display = "none";
                }
                
        }
}


function swapPhoto(photoSRC,theCaption,theCredit,folder) {

if (document.getElementById("caption")) {

	var theImage = document.getElementById("mainPhoto");
	var displayedCaption = document.getElementById("caption");
	var displayedCredit = document.getElementById("credit");
	var imgFolder = folder;
	displayedCaption.firstChild.nodeValue = theCaption;
	displayedCredit.firstChild.nodeValue = theCredit;
	theImage.setAttribute("src", imgFolder+photoSRC);

    }
  }
