function createDOM() {
	var dom = new ActiveXObject("MSXML2.DOMDocument");

	dom.async = false;

	return dom;
}

function createFTDOM() {
	var ftdom = new ActiveXObject("MSXML2.FreeThreadedDOMDocument");

	ftdom.async = false;

	return ftdom;
}

function createXSLTmp(xsldoc) {
	var xsltmp = new ActiveXObject("MSXML2.XSLTemplate");

	xsltmp.stylesheet=xsldoc;

	return xsltmp;
}

