function UpdateLocation(sTheURL) {
	var myframeL = parent.document.getElementsByTagName('iframe')[0];

	if (myframeL !== null)
	{

		var docL = myframeL.contentWindow.document;
		if (docL !== null)
		{

			var elems = docL.getElementsByTagName("a");
			for (var i = elems.length; i--;)
			{
				if (elems[i].innerHTML == sTheURL)
				{
					if (elems[i].style.color != "black")
					{
						elems[i].style.color = "black";
					}
					if (elems[i].style.fontWeight != 600)
					{
						elems[i].style.fontWeight = 600;
					}
				}
				else
				{
					if (elems[i].style.color != "#00709f")
					{
						elems[i].style.color = "#00709f";
					}
					if (elems[i].style.fontWeight != 400)
					{
						elems[i].style.fontWeight = 400;
					}
				}

			}
		}
    }
    return false;
}