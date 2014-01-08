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

function BTN_OnClick(sIdCode, sTooltip)
{
	if (clipboardData && clipboardData.setData)
	{
		clipboardData.setData("Text", document.getElementById(sIdCode).innerText);

		var oToolTip = document.getElementById(sTooltip);
		if (oToolTip != null)
		{
			oToolTip.innerHTML = "<table><tr><td class=tooltip>Copied</td></tr></table>";
			oToolTip.style.pixelLeft = (event.x+20+document.documentElement.scrollLeft);
			oToolTip.style.pixelTop = (event.y+15+document.documentElement.scrollTop);
			oToolTip.style.visibility = "visible";
			setTimeout(function(){oToolTip.style.visibility="hidden";}, 1200)
		}
	}
}

function Btn_OnMouseOver(oItem)
{
	oItem.style.color="#fff";
	oItem.style.background="#999";

}

function Btn_OnMouseOut(oItem, sTooltip)
{
	oItem.style.color = "#444";
	oItem.style.background = "#E9E9E9";

	var oToolTip = document.getElementById(sTooltip);
	if (oToolTip != null)
	{
		oToolTip.style.visibility = "hidden";
	}
}

function MSDN_Nav(sTheURL, sFunc)
{
	if (window.XMLHttpRequest)
	{
		var xmlhttp = new XMLHttpRequest();
	}
	else
	{
		var xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
	}

	if (sFunc.charAt(0) == "_") sFunc = sFunc.slice(1)
	sTheURL_Search = "http://social.msdn.microsoft.com/search/en-US?query="

	xmlhttp.open("GET", sTheURL, true);

	xmlhttp.onreadystatechange = function()
	{
		if (xmlhttp.readyState == 4)
		{
			var regex = new RegExp("<title[^>]*>([^<]+)<\/title>", "i");
			var arr = xmlhttp.responseText.match(regex)
			var title = "";
			try
			{
				title = arr[1];
			}
			catch(e)
			{
			}

			regex = new RegExp(sFunc, "ig");

			if (title.search(regex) == -1)
			{
				window.open(sTheURL_Search + sFunc)
			}
			else
			{
				window.open(sTheURL)
			}
		}
	}

	xmlhttp.send()
}
