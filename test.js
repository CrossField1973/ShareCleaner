function changeP(){
    document.getElementById("numberOfAnime").innerHTML = "Test";
}

function countAnime(){
    var div = document.getElementsByClassName("list-entries")[0]; 
    var num = 0; children = div.childNodes; 
    for(var i = 0; i < children.length; i++){
        if(children[i].nodeName == "A"){
            num++;
        }
    }
    document.getElementById("numberOfAnime").innerHTML = num;
}