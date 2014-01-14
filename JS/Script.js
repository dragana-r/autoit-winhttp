var iDBindex = -1;
var bFirstRun = true;

window.onload = init;
window.onresize = init

var iStartX = 0;
var iStartW = 0;

var searchDB = new Array();
// DEFINE NEW SEARCHABLE ITEMS HERE, FOLLOWING THE FORMAT AS BELOW;
// searchDB[nextnumber] = new searchOption("TitleOfResult", "Description", "ShortDescription", "AltLink");
searchDB[0] = new searchOption("WinHttp Function Notes", "", "WinHttp-UDFs for AutoIt and this helpfile are...", "../CHM_HomePage");
//--PLACEHOLDER-I-DO-NOT-REMOVE-ME--//

function init()
{
    var window_height;
    var window_width

    if (document.documentElement)
    {
        window_height = document.documentElement.offsetHeight;
        window_width = document.documentElement.offsetWidth;
    }
    else if (window.innerHeight)
    {
        window_height = window.innerHeight;
        window_width = window.innerWidth;
    }

    var iFrame = document.getElementById('idFrame')
    var iFramex = document.getElementById('idFramex')

    if (iFrame && iFramex)
    {
        if (window_height > 120)
        {
            iFrame.height = window_height - 120;
            iFramex.height = window_height - 120;
        }
        if (bFirstRun)
        {
            iFrame.width = window_width * 0.2;
            iFramex.width = window_width * 0.8;
        }
    }

    // bFirstRun = false;

    var iInput = document.getElementById('in')
    if (iInput)
    {
        iInput.style.color = "gray";
        iInput.value = "Search";
		iInput.disabled = false;
    }

    setTimeout(LoadSearchKeywords, 100); // give it time to load

}

function ResizeX(oItem, oItem2)
{

    if (iStartX == 0) return;

    var iWidth = iStartW + (event.x - iStartX);

    if (iWidth > 0)
    {
        var window_width;
        if (document.documentElement)
        {
            window_width = document.documentElement.offsetWidth;
        }
        else if (window.innerWidth)
        {
            window_width = window.innerWidth;
        }

        var iDelta = iWidth + oItem.width
        if (oItem)
        {
            oItem.width = iWidth;
        }
        if (oItem2 && window_width > iWidth)
        {
            oItem2.width = window_width - iWidth;
        }

    }
}

function ExitResizeX(oItem)
{
    if (oItem == null)
    {
        oItem = document.getElementById('rsz');
    }
    oItem.releaseCapture();
    iStartX = 0;
}

function EnterResizeX(oItem, oItem2)
{
    oItem.setCapture();
    iStartX = event.x;
    iStartW = oItem2.offsetWidth;
    if (document.getElementById('rightframe').src.substring(0, 4) == "http") // external page loaded into the frame
    {
        setTimeout(ExitResizeX, 500);
    }
}

function SetIFrameSource(url, elem)
{
    var myframe = parent.document.getElementsByTagName('iframe')[1];
    if (myframe !== null)
    {
        if (myframe.src)
        {
            myframe.src = url
        }
        else if (myframe.contentWindow !== null && myframe.contentWindow.location !== null)
        {
            myframe.contentWindow.location = url
        }
        else
        {
            myframe.setAttribute('src', url)
        }
    }

    {
        var elems = document.getElementsByTagName("a");
        for (var i = elems.length; i--;)
        {
            elems[i].style.color = "#00709f";
            elems[i].style.fontWeight = 400;
        }
        if (elem != null)
        {
            elem.style.color = "black";
            elem.style.fontWeight = 600
        }
    }

    return false;
}


function UpdateLocation(sTheURL)
{
    var myframeL = document.getElementById('leftframe')
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


function UpdateLocationByRightTitle()
{
    var myframeL = document.getElementById('leftframe')
    if (myframeL !== null)
    {

        var myframeR = document.getElementById('rightframe')
        if (myframeR !== null)
        {
            try // access may be denied
            {
                var windowR = myframeR.contentWindow
                var docR = windowR.document;
            }
            catch (e)
            {
                return false;
            }

            if (docR !== null)
            {
                var sTheURL = docR.title

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
        }
    }
    return false;
}


function GoBack()
{
    history.back();
    setTimeout(UpdateLocationByRightTitle, 100); // give it time to load
    return false;
}


function GoForward()
{
    history.forward();
    setTimeout(UpdateLocationByRightTitle, 100); // give it time to load
    return false;
}

function GoHome()
{
    window.open("//--PLACEHOLDER-II-DO-NOT-REMOVE-ME--//");
}


function LoadSearchKeywords(index)
{
    if (index == null)
    {
        index = 0;
    }

    var myframe = document.getElementById('rightframe')

    if (myframe !== null)
    {

        var doc = myframe.contentWindow.document;
        if (doc)
        {
            bodytext = "" + doc.body.innerText
            bodytext = bodytext.replace(/\t/g, " ")
            bodytext = bodytext.replace(/\n/g, " ")
            bodytext = bodytext.replace(/\r/g, " ")
            bodytext = bodytext.replace(/ +/g, " ")
            searchDB[index].description = bodytext
        }
    }
}


function DoSearch()
{
    var myframe = document.getElementsByTagName('iframe')[1];
    if (myframe)
    {
        try // access may be denied
        {
            var windowX = myframe.contentWindow
            var doc = windowX.document;
        }
        catch (e)
        {
            document.getElementById("in").disabled = true;
			return false;
        }

        if (doc)
        {
            // Do the search
            var sTerm = document.getElementById("in").value;

            if (sTerm.toLowerCase() == "tetris")
            {
                SetIFrameSource("tetris.htm");
                setTimeout(UpdateLocationByRightTitle, 50); // give it time to load
                return false;
            }
            else if (doc.title == "Tetris")
            {
                history.back();
                setTimeout(DoSearch, 100);
                return false;
            }

            results = performSearch(sTerm);
            var regex = new RegExp(sTerm, "ig");
            var sResult = "";

            if (results)
            {
                // This means that there are search results to be displayed.
                // Loop through them and make it pretty! :)
                if (is_array(results))
                {
                    sResult += "<h2>Search Results for \"" + sTerm + "\" (" + results.length + ")</h2>";
                    sResult += "<ol class=\"result\">";
                    for (r = 0; r < results.length; r++)
                    {
                        result = searchDB[results[r]];

                        /////////////////////////////////////////////////////////////
                        // This is where you modify the formatting of the results
                        sResult += "<li class=\"result\"><div class=\"result-title\"><a href=\"#\" onclick='parent.document.getElementsByTagName(\"iframe\")[1].src = \"" + (result.altlink ? result.altlink : result.heading) + ".htm\"; parent.UpdateLocation(\"" + result.heading + "\"); return false'>" + result.heading + "</a></div>";
                        sResult += "<div class=\"result\">" + (result.shortdesc ? result.shortdesc : result.description).replace(regex, "<strong>$&</strong>") + "</div></li>";
                        /////////////////////////////////////////////////////////////
                    }
                    sResult += "</ol>";
                }
                    // If it's not an array, then we got an error message, so display that
                    // rather than results
                else
                {
                    sResult += "<i>" + results + "</i>";
                }

            }
            doc.body.innerHTML = sResult;
        }
    }
    return false;
}

// These are the available "error strings" you can change them to affect the output
// of the search engine.
ERR_NoOptions = "You didn't specify where to search for your keywords, please try again.";
ERR_NoSearchTerms = "You didn't enter any words to search for, please enter some words to search for and try again.";
ERR_NoResults = "Your search found no results, try less specific terms.";

// Performs an actual search and then returns the index number(s) in the db array
// where it found this element.
// keywords = the string they searched for (each space-separated word
//     is searched for separately)
// options can be
//     1 = search keywords, not description, not heading
//     2 = search keywords, search description, not heading
//     3 = search all
function performSearch(keywords)
{
    // Check to make sure they entered some search terms
    if (!keywords || keywords.length == 0)
    {
        return ERR_NoSearchTerms;
    }

    searchDescription = true;
    searchHeading = true

    // Setting up the keywords array for searching
    // Remove common punctuation
    keywords = keywords.replace("\.,'", "");

    // get them all into an array so we can loop thru them
    // we assume a space was used to separate the terms
    searchFor = keywords.split(" ");

    // This is where we will be putting the results.
    results = new Array();

    // Loop through the db for potential results
    // For every entry in the "database"
    for (sDB = 0; sDB < searchDB.length; sDB++)
    {

        // For every search term we are working with
        for (t = 0; t < searchFor.length; t++)
        {
            // Check in the heading for the term if required
            if (searchHeading)
            {
                if (searchDB[sDB].heading.toLowerCase().indexOf(searchFor[t].toLowerCase()) != -1)
                {
                    if (!in_array(String(sDB), results))
                    {
                        results[results.length] = String(sDB);
                    }
                }
            }

            // Check in the description for the term if required
            if (searchDescription)
            {
                if (searchDB[sDB].description.toLowerCase().indexOf(searchFor[t].toLowerCase()) != -1)
                {
                    if (!in_array(String(sDB), results))
                    {
                        results[results.length] = String(sDB);
                    }
                }
            }
        }
    }

    if (results.length > 0)
    {
        return results;
    }
    else
    {
        return ERR_NoResults;
    }
}

// Constructor for each search engine item.
// Used to create a record in the searchable "database"
function searchOption(heading, description, shortdesc, altlink)
{
    this.heading = heading;
    this.description = description;
    this.shortdesc = "";
    if (shortdesc != null)
    {
        this.shortdesc = shortdesc;
    }
    this.altlink = "";
    if (altlink != null)
    {
        this.altlink = altlink;
    }
    return this;
}

// Returns true or false based on whether the specified string is found
// in the array.
// This is based on the PHP function of the same name.
// stringToSearch = the string to look for
// arrayToSearch  = the array to look for the string in.
function in_array(stringToSearch, arrayToSearch)
{
    for (s = 0; s < arrayToSearch.length; s++)
    {
        if (arrayToSearch[s].indexOf(stringToSearch) != -1)
        {
            return true;
            exit;
        }
    }
    return false;
}

// Code from http://www.optimalworks.net/blog/2007/web-development/javascript/array-detection
function is_array(array) { return !(!array || (!array.length || array.length == 0) || typeof array !== 'object' || !array.constructor || array.nodeType || array.item); }


