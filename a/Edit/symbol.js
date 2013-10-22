function symbol_cc(cardID){
	var obj;
	for (var i=1;i<7;i++){
		obj=$id("card"+i);
		obj.style.fontWeight="normal";
		obj.className="clicktext";
	}
	obj=$id("card"+cardID);
	obj.className="grayfont";

	for (var i=1;i<7;i++){
		obj=$id("content"+i);
		obj.style.display="none";
	}
	obj=$id("content"+cardID);
	obj.style.display="";
}
symbol_cc(1);