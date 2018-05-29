    var docRef = app.activeDocument;
    var docRefAll = docRef.pageItems.length;  
    var sizeIndex=0.352777777777778;
    var shift=100;
    readXML();
    shirina=parseFloat(shirina);
    vysota=parseFloat(vysota);
//    alert (vysota);
    chngArt(0, shirina, vysota);
    chngArt(1, 210, 297);

    docRef.artboards.setActiveArtboardIndex(1);

    docRef.swatches.removeAll();

    var xmldata='';

createTables(5+shirina+shift);
createTextus(shirina+shift);
createPodpis();


createIzdelie(izdelie);
createRazmer();

//alert (xmldata);


function createRazmer(){
if (vysota<600){
 razmX = 0;
 }
else {
razmX = 20/sizeIndex;
}

pointhX1=-20/sizeIndex;
pointhY1=0;
pointhY2=-(vysota/2-20)/sizeIndex;
pointhY3=-(vysota/2+20)/sizeIndex;
pointhY4=-vysota/sizeIndex;
createVertical(pointhX1, pointhY1, pointhY2+razmX, 1, vysota);
createVertical(pointhX1, pointhY3-razmX, pointhY4, 2, vysota);
if (vysota==vysota.toFixed()){
 addRazmer(-20/sizeIndex,-(vysota/2)/sizeIndex,vysota+".00",90);
}
else {
 addRazmer(-20/sizeIndex,-(vysota/2)/sizeIndex,vysota,90);    
}
pointhX1=0;
pointhY1=-(vysota+20)/sizeIndex;
pointhX2=(shirina/2-20)/sizeIndex;
pointhX3=(shirina/2+20)/sizeIndex;
pointhX4=shirina/sizeIndex;

if (dno!='' && dno!=0){
pointhY1 -= dno/sizeIndex;
}

createHorizontal(pointhX1, pointhX2-razmX, pointhY1, 1, shirina);
createHorizontal(pointhX3+razmX, pointhX4, pointhY1, 2, shirina);
if (shirina==shirina.toFixed()){
 addRazmer(shirina/2/sizeIndex,pointhY1,shirina+".00",0);
}
else {
    addRazmer(shirina/2/sizeIndex,pointhY1,shirina,0);
    }
if (dno!='' && dno!=0){

pointhX1=-30/sizeIndex;
pointhY1=-vysota/sizeIndex;
pointhY2=-(vysota+dno)/sizeIndex;
createVertical(pointhX1, pointhY1, pointhY2, 1, vysota);
createVertical(pointhX1, pointhY1, pointhY2, 2, vysota);
 addRazmer(pointhX1,pointhY1+(dno+10)/sizeIndex,"донный подворот "+dno+".00",90);


}


}

function createHorizontal(x,x2,y,corner,text){
 tmpg = docRef.groupItems.add();
 line = tmpg.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[x,y],[x2,y]]);
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;
 if (corner==0){
 }
 else if (corner==1){
 line = tmpg.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[x,y-5/sizeIndex],[x,y+10/sizeIndex]]);
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;

 lineList=[[x,y],[x+10/sizeIndex,y-2/sizeIndex],[x+10/sizeIndex,y+2/sizeIndex],[x,y]];
 newPath = tmpg.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.closed=true;
 newPath.fillColor.black=100;
 }
 else{
 line = tmpg.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[x2,y-5/sizeIndex],[x2,y+10/sizeIndex]]);
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;


 lineList=[[x2,y],[x2-10/sizeIndex,y-2/sizeIndex],[x2-10/sizeIndex,y+2/sizeIndex],[x2,y]];
 newPath = tmpg.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.closed=true;
 newPath.fillColor.black=100;

 }

}


function createVertical(x,y,y2,corner, text){
 tmpg = docRef.groupItems.add()
 line = tmpg.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[x,y],[x,y2]]);
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;
 if (corner==0){
 }
 else if (corner==1){
 line = tmpg.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[x-5/sizeIndex,y],[x+10/sizeIndex,y]]);
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;

 lineList=[[x,y],[x+2/sizeIndex,y-10/sizeIndex],[x-2/sizeIndex,y-10/sizeIndex],[x,y]];
 newPath = tmpg.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.closed=true;
 newPath.fillColor.black=100;
 }
 else{
 line = tmpg.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[x-5/sizeIndex,y2],[x+10/sizeIndex,y2]]);
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;


 lineList=[[x,y2],[x+2/sizeIndex,y2+10/sizeIndex],[x-2/sizeIndex,y2+10/sizeIndex],[x,y2]];
 newPath = tmpg.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.closed=true;
 newPath.fillColor.black=100;
 }
}
function addRazmer(x,y,text,angle,txtSize){
  if (!txtSize){
    if (vysota<600){
  txtSize=21;
  }
  else {
  txtSize=40;
  }
  }
    newTF=docRef.textFrames.add();
    //newTF.textRange.characterAttributes.TextFont='Myriad Pro';
    newTF.textRange.characterAttributes.language=LanguageType.RUSSIAN;
    newTF.contents=text+' mm';
    newTF.textRange.characterAttributes.size = txtSize;
    newTF.rotate(angle);
    newTF.textRange.paragraphAttributes.justification = Justification.CENTER;
    newTF.left=x-newTF.width/2;
    newTF.top=y+newTF.height/2;

}

function createPodpis(){
     txtX=(42+shirina+shift)/sizeIndex;
    txtY=-230.5/sizeIndex;
addText(txtX,txtY,zakazchik,0,9);
     txtX=(42+shirina+shift)/sizeIndex;
    txtY=-234.5/sizeIndex;
addText(txtX,txtY,izdelie+', "'+nazvanie+'"',0,9);
     txtX=(48.5+shirina+shift)/sizeIndex;
    txtY=-238.5/sizeIndex;
addText(txtX,txtY,formula,1,9); 
   day=new Date().getDate();
  month=new Date().getMonth();
  month++;
  month=(month<10?"0"+month: month);
  year=new Date().getFullYear();
  data=day+"."+month+"."+year+"г";
    txtX=(48.5+shirina+shift)/sizeIndex;
    txtY=-254.5/sizeIndex;
addText(txtX,txtY,data,1,9);   
     txtX=(42+shirina+shift)/sizeIndex;
    txtY=-258.5/sizeIndex;
addText(txtX,txtY,manager,0,9);
     txtX=(42+shirina+shift)/sizeIndex;
    txtY=-262.5/sizeIndex;
addText(txtX,txtY,"Солодкий Е. А. /",0,9);

     txtX=(101+shirina+shift)/sizeIndex;
    txtY=-238.5/sizeIndex;
addText(txtX,txtY,printer,1,9);
     txtX=(101+shirina+shift)/sizeIndex;
    txtY=-242.5/sizeIndex;
addText(txtX,txtY,raport,1,9);
     txtX=(101+shirina+shift)/sizeIndex;
    txtY=-246.5/sizeIndex;
  if (printer=='8x'){
   klishe='1,7';
   }
   else{
   klishe='2,84';
   }
xmldata+=klishe+';';
addText(txtX,txtY,klishe,1,9);

     txtX=(48.5+shirina+shift)/sizeIndex;
    txtY=-246.5/sizeIndex;
addText(txtX,txtY,material,1,8);

     txtX=(101+shirina+shift)/sizeIndex;
    txtY=-250.5/sizeIndex;
addText(txtX,txtY,technomer,1,8);
     txtX=(101+shirina+shift)/sizeIndex;
    txtY=-254.5/sizeIndex;
    if (izdelie=='банан'){
      kadr=parseInt(raport)/parseInt(shirina);
    }
    else{
  kadr=parseInt(raport)/parseInt(vysota);
  }
kadr=(kadr<1?1: kadr);
if (printer=='8x'){
	kadr="HD150+FTD";
}
else {
	kadr="Regular";	
}
addText(txtX,txtY,kadr,1,9);
     txtX=(127.5+shirina+shift)/sizeIndex;
    txtY=-230.5/sizeIndex;
    if (izdelie=='ламинат' || zakazchik == 'Гранд Флекс' || zakazchik == 'ГрандФлекс'){
    addText(txtX,txtY,"Печать обратная",1,9);
xmldata+='обратная;'
    }
    else{
addText(txtX,txtY,"Печать прямая",1,9);
xmldata+='прямая;'
}
     txtX=(111+shirina+shift)/sizeIndex;
    txtY=-234.5/sizeIndex;
  if (printer=='8x'){
  addText(txtX,txtY,"Цветопроба      ДА    НЕТ",0,9);
xmldata+='Изготовление цветопробы';

 tmpg11 = docRef.groupItems.add();
 line11 = tmpg11.pathItems.add();
 line11.stroked = true;
 line11.setEntirePath([[txtX+55,txtY-0.5/sizeIndex],[txtX+55,txtY-4.5/sizeIndex]]);
 line11.strokeWidth=0.2/sizeIndex;
 line11.strokeColor.black=60;

 line11 = tmpg11.pathItems.add();
 line11.stroked = true;
 line11.setEntirePath([[txtX+75,txtY-0.5/sizeIndex],[txtX+75,txtY-4.5/sizeIndex]]);
 line11.strokeWidth=0.2/sizeIndex;
 line11.strokeColor.black=60;

   }
   else{
//addText(txtX,txtY,"CtP",1,9);
xmldata+='';
   }
if (izdelie=='рукав' || izdelie=='полурукав' || izdelie=='полотно' || izdelie=='ламинат'){
     txtX=(127.5+shirina+shift)/sizeIndex;
    txtY=-238/sizeIndex;
addText(txtX,txtY,"Схема размотки",1,8);
     txtX=(127.5+shirina+shift)/sizeIndex;
     txtY=-240.5/sizeIndex;
addText(txtX,txtY,"готовой продукции",1,8);
//alert (namotk);
namot(namotk,shift);
}

  txtX=(5+shirina+shift)/sizeIndex;
    txtY=-267.5/sizeIndex;
addText(txtX,txtY,"Данный документ не является цветопробой. Цветопередача при печати будет иной.",0,9);
  txtX=(5+shirina+shift)/sizeIndex;
    txtY=-270.5/sizeIndex;
addText(txtX,txtY,"После подписания макета вся ответственность по текстовой и графической информации,",0,9);
  txtX=(5+shirina+shift)/sizeIndex;
  txtY=-273.5/sizeIndex;
addText(txtX,txtY,"а также их наличие в подписываемом макете возлагается на Заказчика!",0,9);  
  txtX=(5+shirina+shift)/sizeIndex;
    txtY=-276.5/sizeIndex;
addText(txtX,txtY,"Во избежание ошибок проверяйте текст и размеры перед подтверждением.",0,9);  
  txtX=(5+shirina+shift)/sizeIndex;
    txtY=-279.5/sizeIndex;
addText(txtX,txtY,"Утвержденным макет считается только при наличии Вашей подписи и печати.",0,9); 
   txtX=(5+shirina+shift)/sizeIndex;
   txtY=-282.5/sizeIndex;
addText(txtX,txtY,"Рабочие файлы являются интеллектуальной собственностью Исполнителя и не предоставляются Заказчику.",0,9); 
  txtX=(85+shirina+shift)/sizeIndex;
    txtY=-287.5/sizeIndex;
addText(txtX,txtY,"Макет со стороны заказчика утверждаю: ______________________",0,11.5);   
}
  
function createTextus(pos){
textus=["заказчик","название изделия","формула печати","цвет материала","материал","ширина к запечатке","дата","отв. менеджер","нач. производства"];
shiftY=[0,4,8,12,16,20,24,28,32]
for (x=0; x<shiftY.length; x++){
posX=(pos+22)/sizeIndex;
posY=(-230.5-shiftY[x])/sizeIndex;
addText(posX,posY,textus[x],1,9);
}
textus=["печатная машина","печатный вал","толщина формы, мм","технологический номер","тип растрирования"];
shiftY=[8,12,16,20,24]
for (x=0; x<shiftY.length; x++){
posX=(pos+75)/sizeIndex;
posY=(-230.5-shiftY[x])/sizeIndex;
addText(posX,posY,textus[x],1,9);
}

textus={"Праймер":"Праймер","C":"cyan","M":"magenta","Y":"yellow","K":"black","белый":"white","white":"white","White":"white","485C":"485", "P":"Pantone", "красный":"485", "синий":"Reflex Blue", "зеленый":"340", "зелёный":"340", "оранжевый":"021", "оранж":"021", "серебро":"Silver", "фиолетовый":"Violet","золото":"873","салатовый":"360","коричневый":"483","черный":"black","Reflexblue":"Reflex Blue","малиновый":"magenta","серебро":"877"};
textus2=["C","M","Y","K","белый","485", "P", "красный", "синий"];

shiftY=[0,4,8,12,16,20,24,28]                                          
for (x=0; x<allColors.length; x++){
posX=(pos+158.5)/sizeIndex;
posY=(-234.5-shiftY[x])/sizeIndex;

if (allColors[x]!='C' && allColors[x]!='M' && allColors[x]!='Y' && allColors[x]!='K' && allColors[x]!='белый' && allColors[x]!='white' && allColors[x]!='Праймер' && isNaN(parseInt(allColors[x]))){
//allColors[x]='P';
}
if (isNaN(parseInt(allColors[x]))){
curCol=allColors[x];
curoCol=textus[curCol];
if (curoCol==undefined){
curoCol='P.';
}
if (curoCol!="cyan" && curoCol!="magenta" && curoCol!="yellow" && curoCol!="black"){
addColor(curoCol);
addText(posX,posY,curoCol,0,9);
}
else {
addText(posX,posY,curoCol,0,9);
}
}
else{
addText(posX,posY,"Pantone "+parseInt(allColors[x]),0,9);
addColor(allColors[x]);

}

}

shiftY=0;
for (x=1; x<allColors.length+1; x++){
posX=(pos+148.5)/sizeIndex;
posY=(-234.5-shiftY)/sizeIndex;
addText(posX,posY,x,1,9);
shiftY+=4;
}

textus=["124","124","124","124","---","---","---","---"];
textus2=["7.5","67.5","82.5","37.5","0","0","0","0"];
shiftY=[0,4,8,12,16,20,24,28]
for (x=0; x<allColors.length; x++){
posX=(pos+187.5)/sizeIndex;
posX2=(pos+200)/sizeIndex;
posY=(-234.5-shiftY[x])/sizeIndex;
addText(posX,posY,textus[x],1,9);
addText(posX2,posY,textus2[x],1,9);
}

textus=["№","цвет","наименование","lpi",];
shiftX=[0,6,14.5,18.5]
shiftus=0;
for (x=0; x<shiftX.length; x++){
shiftus+=shiftX[x];
posX=(pos+148.5+shiftus)/sizeIndex;
posY=(-230.5)/sizeIndex;
pantone=textus[x];
addText(posX,posY,pantone,1,9);

}

posX=(pos+187.5)/sizeIndex;
posY=(-226)/sizeIndex;
addText(posX,posY,'линиатура',1,8.5);

posX=(pos+200)/sizeIndex;
posY=(-226)/sizeIndex;
addText(posX,posY,'угол',1,8.5);


posX=(pos+150.5)/sizeIndex;
posY=(-274)/sizeIndex;
addText(posX,posY,'!',1,30);

posX=(pos+153)/sizeIndex;
posY=(-268)/sizeIndex;
addText(posX,posY,'Готовое изделие может содержать ',0,9);
posX=(pos+153)/sizeIndex;
posY=(-271)/sizeIndex;
addText(posX,posY,'технологические метки: микроточки,',0,9);
posX=(pos+153)/sizeIndex;
posY=(-274)/sizeIndex;
addText(posX,posY,'приладочные кресты, подписи и др.',0,9);




}

function createIzdelie(izdelie){
if (izdelie=='банан'){
 itRef = docRef.pathItems.rectangle(0, 0, shirina/sizeIndex, vysota/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;
 itRef.fillColor.cyan=0;
 posX=(shirina/2-40)/sizeIndex;
 itRef = docRef.pathItems.roundedRectangle(-50/sizeIndex, posX, 80/sizeIndex, 20/sizeIndex, 10/sizeIndex, 10/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;
}
else if (izdelie=='банан с усилением'){
 itRef = docRef.pathItems.rectangle(0, 0, shirina/sizeIndex, vysota/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;
 itRef.fillColor.cyan=0;
 posX=(shirina/2-40)/sizeIndex;
 itRef = docRef.pathItems.rectangle(-20/sizeIndex, posX-35/sizeIndex, 150/sizeIndex, 80/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;
 itRef.strokeDashes=[5/sizeIndex,5/sizeIndex];
 itRef = docRef.pathItems.roundedRectangle(-50/sizeIndex, posX, 80/sizeIndex, 20/sizeIndex, 10/sizeIndex, 10/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;

}
else if (izdelie=='майка'){
 point1=(shirina/2-nozh*5)/sizeIndex;
 point2=point1+(nozh*10/sizeIndex);
 lineList=[[0,0],[point1,0],[point1,-150/sizeIndex],[point2,-150/sizeIndex],[point2,0],[shirina/sizeIndex,0],[shirina/sizeIndex,-vysota/sizeIndex],[0,-vysota/sizeIndex],[0,0]];
 newPath = app.activeDocument.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.strokeWidth=1/sizeIndex;
 newPath.strokeColor.black=100;
 newPath.closed=true;
 if (bokovye<11 && bokovye>0){
 bokovye=bokovye*10;
 }
 line = docRef.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[bokovye/sizeIndex, -150/sizeIndex],[bokovye/sizeIndex,-vysota/sizeIndex]]);
 line.strokeWidth=0.75/sizeIndex;
 line.strokeColor.black=100;
 line.strokeDashes=[5/sizeIndex,5/sizeIndex];

 line = docRef.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[shirina/sizeIndex-bokovye/sizeIndex, -150/sizeIndex],[shirina/sizeIndex-bokovye/sizeIndex,-vysota/sizeIndex]]);
 line.strokeWidth=0.75/sizeIndex;
 line.strokeColor.black=100;
 line.strokeDashes=[5/sizeIndex,5/sizeIndex];

pointhX1=0;
pointhY1=-(170)/sizeIndex;
pointhX2=(bokovye/2-20)/sizeIndex;
pointhX3=(bokovye/2+20)/sizeIndex;
pointhX4=bokovye/sizeIndex;
createHorizontal(pointhX1, pointhX2, pointhY1, 1, bokovye);
createHorizontal(pointhX3, pointhX4, pointhY1, 2, bokovye);
addRazmer(bokovye/2/sizeIndex,pointhY1,bokovye+".00",0);

shift=(shirina-bokovye)/sizeIndex;
pointhX1+=shift;
pointhX2+=shift;
pointhX3+=shift;
pointhX4+=shift;
createHorizontal(pointhX1, pointhX2, pointhY1, 1, bokovye);
createHorizontal(pointhX3, pointhX4, pointhY1, 2, bokovye);
addRazmer((shirina-(bokovye/2))/sizeIndex,pointhY1,bokovye+".00",0);

 }
else {
 itRef = docRef.pathItems.rectangle(0, 0, shirina/sizeIndex, vysota/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;
 itRef.fillColor.cyan=0;
}

if (izdelie=='петлевая'){
    
 petlya =docRef.groupItems.add();
 
 point1=(shirina/2-70)/sizeIndex;
 point2=point1+(30/sizeIndex);
 lineList=[[point1,0],[point2,0],[point2,80/sizeIndex],[point1,50/sizeIndex],[point1,0]];
 newPath = petlya.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.strokeWidth=1/sizeIndex;
 newPath.strokeColor.black=100;
 newPath.closed=true;

 point1=(shirina/2+40)/sizeIndex;
 point2=point1+(30/sizeIndex);
 lineList=[[point1,0],[point2,0],[point2,50/sizeIndex],[point1,80/sizeIndex],[point1,0]];
 newPath = petlya.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.strokeWidth=1/sizeIndex;
 newPath.strokeColor.black=100;
 newPath.closed=true;
 
 point1=(shirina/2-40)/sizeIndex;
 point2=point1+(80/sizeIndex);
 lineList=[[point1,50/sizeIndex],[point2,50/sizeIndex],[point2,80/sizeIndex],[point1,80/sizeIndex],[point1,50/sizeIndex]];
 newPath = petlya.pathItems.add();
 newPath.setEntirePath(lineList);
 newPath.strokeWidth=1/sizeIndex;
 newPath.strokeColor.black=100;
 newPath.closed=true;
}
if (dno!='' && dno!=0){
 itRef = docRef.pathItems.rectangle(-vysota/sizeIndex,0, shirina/sizeIndex, dno/sizeIndex);
 itRef.strokeWidth=1/sizeIndex;
 itRef.strokeColor.black=100;
 itRef.fillColor.cyan=0;
}
if (bokovye!=0 && izdelie!='майка'){
 if (bokovye<11 && bokovye>0){
 bokovye=bokovye*10;
 }
 line = docRef.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[bokovye/sizeIndex, -0/sizeIndex],[bokovye/sizeIndex,-vysota/sizeIndex]]);
 line.strokeWidth=0.75/sizeIndex;
 line.strokeColor.black=100;
 line.strokeDashes=[5/sizeIndex,5/sizeIndex];

 line = docRef.pathItems.add();
 line.stroked = true;
 line.setEntirePath([[shirina/sizeIndex-bokovye/sizeIndex, -0/sizeIndex],[shirina/sizeIndex-bokovye/sizeIndex,-vysota/sizeIndex]]);
 line.strokeWidth=0.75/sizeIndex;
 line.strokeColor.black=100;
 line.strokeDashes=[5/sizeIndex,5/sizeIndex];

pointhX1=0;
pointhY1=-(10)/sizeIndex;
pointhX2=(bokovye/2-20)/sizeIndex;
pointhX3=(bokovye/2+20)/sizeIndex;
pointhX4=bokovye/sizeIndex;
createHorizontal(pointhX1, pointhX2, pointhY1, 1, bokovye);
createHorizontal(pointhX3, pointhX4, pointhY1, 2, bokovye);
addRazmer(bokovye/2/sizeIndex,pointhY1,bokovye+".00",0);

shift=(shirina-bokovye)/sizeIndex;
pointhX1+=shift;
pointhX2+=shift;
pointhX3+=shift;
pointhX4+=shift;
createHorizontal(pointhX1, pointhX2, pointhY1, 1, bokovye);
createHorizontal(pointhX3, pointhX4, pointhY1, 2, bokovye);
addRazmer((shirina-(bokovye/2))/sizeIndex,pointhY1,bokovye+".00",0);
}
}


function createTables(x){
rectS=[35,70,35, 35,70,35, 35,17.5,35,17.5,35, 35,17.5,35,17.5, 35,17.5,35,17.5, 35,17.5,35,17.5, 35,17.5,35,17.5, 35,70, 35,70]; 
rectW=[4,4,4, 4,4,4, 4,4,4,4,28, 4,4,4,4, 4,4,4,4, 4,4,4,4, 4,4,4,4, 4,4, 4,4];
rectZ=[0,0,1, 0,0,1, 0,0,0,0,1,  0,0,0,1, 0,0,0,1, 0,0,0,1, 0,0,0,1, 0,1, 0,1];
posX=x/sizeIndex;
posY=-231/sizeIndex;
shiftX=0;
shiftY=0;
for (i=0; i<rectW.length; i++){
itRef = docRef.pathItems.rectangle(posY+shiftY, posX+shiftX, rectS[i]/sizeIndex, rectW[i]/sizeIndex);
if (rectZ[i]>0){
shiftY-=4/sizeIndex;
shiftX=0;
}
else{
shiftX+=rectS[i]/sizeIndex;
}

noColor = new NoColor();
itRef.fillColor = noColor;
itRef.strokeWidth=0.2/sizeIndex;
itRef.strokeColor.black=60;
}

itRef = docRef.pathItems.rectangle(-226/sizeIndex, (x+175.112)/sizeIndex, 15/sizeIndex, 5/sizeIndex);
itRef.fillColor = noColor;
itRef.strokeWidth=0.2/sizeIndex;
itRef.strokeColor.black=60;

itRef = docRef.pathItems.rectangle(-226/sizeIndex, (x+190.112)/sizeIndex, 10/sizeIndex, 5/sizeIndex);
itRef.fillColor = noColor;
itRef.strokeWidth=0.2/sizeIndex;
itRef.strokeColor.black=60;

posY=-231/sizeIndex;
br=0;
for (p=0; p<allColors.length+1; p++){
rectS=[5,7,22,15,10]; 
posX=x/sizeIndex;

shiftX=400;
shiftY=0;
inkC=[100,0,0,0,0,0,0,0];
inkM=[0,100,0,0,0,0,0,0];
inkY=[0,0,100,0,0,0,0,0];
inkK=[0,0,0,100,0,0,0,0];


for (i=0; i<rectS.length; i++){
itRef = docRef.pathItems.rectangle(posY+shiftY, posX+shiftX, rectS[i]/sizeIndex, 4/sizeIndex);
shiftX+=rectS[i]/sizeIndex;
noColor = new NoColor();

if (p>0 && i==1){
if (allColors[br]=='c' || allColors[br]=='C'){
itRef.fillColor.cyan = 100;
}
if (allColors[br]=='m' || allColors[br]=='M'){
itRef.fillColor.magenta = 100;
}
if (allColors[br]=='y' || allColors[br]=='Y'){
itRef.fillColor.yellow = 100;
}
if (allColors[br]=='k' || allColors[br]=='K'){
itRef.fillColor.black = 100;
}
br++;
}
else {
itRef.fillColor = noColor;
//itRef.fillColor = docRef.spots[i].color;
}

itRef.strokeWidth=0.2/sizeIndex;
itRef.strokeColor.black=60;
}
posY-=4/sizeIndex;

}


}

function namot(type,shift){
 namotka =docRef.groupItems.add();
 namotka.name="namotka";
 line = namotka.pathItems.add();
 line.stroked = true;
 line.filled = false;
if (type[0]=='A'){
 line.setEntirePath([[131.189/sizeIndex,-257.104/sizeIndex],[140.678/sizeIndex,-245.46/sizeIndex],[124.724/sizeIndex,-245.46/sizeIndex],[114.218/sizeIndex,-258.062/sizeIndex],[113.529/sizeIndex,-261.384/sizeIndex],[116.373/sizeIndex,-266.024/sizeIndex],[132.267/sizeIndex,-266.024/sizeIndex]]);
    newPoint = line.pathPoints[4];
    newPoint.anchor = Array(113.529/sizeIndex,-261.384/sizeIndex); 
    newPoint.leftDirection = Array(113.529/sizeIndex,-258.884/sizeIndex);
    newPoint.rightDirection = Array(113.529/sizeIndex,-265.884/sizeIndex);
    shifttxt=0;
}
else {
 line2 = namotka.pathItems.add();
 line2.stroked = true;
 line2.filled = false;
    line2.setEntirePath([[117.218/sizeIndex,-256.323/sizeIndex],[133.267/sizeIndex,-256.323/sizeIndex]]);
    line.setEntirePath([[135.262/sizeIndex,-263.852/sizeIndex],[140.678/sizeIndex,-245.46/sizeIndex],[121.724/sizeIndex,-245.46/sizeIndex],[117.218/sizeIndex,-256.323/sizeIndex],[113.529/sizeIndex,-261.384/sizeIndex],[116.218/sizeIndex,-266.024/sizeIndex],[133.267/sizeIndex,-266.024/sizeIndex]]);
    newPoint = line.pathPoints[4];
    newPoint.anchor = Array(113.529/sizeIndex,-261.384/sizeIndex); 
    newPoint.leftDirection = Array(113.529/sizeIndex,-256.884/sizeIndex);
    newPoint.rightDirection = Array(113.529/sizeIndex,-265.884/sizeIndex);
    shifttxt=2/sizeIndex;
}

 newPoint.pointType = PointType.SMOOTH;
 line.strokeWidth=0.5/sizeIndex;
 line.strokeColor.black=100;

 ellipse = namotka.pathItems.ellipse(-256.356/sizeIndex,129.902/sizeIndex,5.866/sizeIndex,9.698/sizeIndex);
 ellipse.stroked = true;
 ellipse.filled = false;
 ellipse2 = namotka.pathItems.ellipse(-258.979/sizeIndex,131.489/sizeIndex,2.693/sizeIndex,4.452/sizeIndex);
 ellipse2.stroked = true;
 ellipse2.filled = false;
 ellipse3 = namotka.pathItems.ellipse(-260.216/sizeIndex,132.237/sizeIndex,1.197/sizeIndex,1.978/sizeIndex);
 ellipse3.filled=true;
 ellipse3.stroked = false;
 ellipse3.fillColor.cyan=0;
 ellipse3.fillColor.black=100;
 ellipse.strokeWidth=0.5/sizeIndex;
 ellipse2.strokeWidth=0.5/sizeIndex;

    newTF=namotka.textFrames.add();
    newTF.contents=type;
    newTF.left=119.837/sizeIndex;
    newTF.top=-260.038/sizeIndex;
    newTF.textRange.characterAttributes.size = 13;

    newTF=namotka.textFrames.add();
    newTF.contents="PRINT";
    newTF.left=121.9/sizeIndex+shifttxt;
    newTF.top=-250/sizeIndex;
    newTF.textRange.characterAttributes.size = 10;
if (type=='A2' || type=='B2'){
    newTF.rotate(180);
    }
if (type=='A3' || type=='B3'){
    newTF.rotate(-120);
    }
if (type=='A4' || type=='B4'){
    newTF.rotate(60);
    }


    shiftik=((docRef.artboards[0].artboardRect[2])*sizeIndex+shift*sizeIndex+190*sizeIndex);
namotka.left = namotka.left+shiftik/sizeIndex;

 }

function addText(x,y,text,forw,size){
    newTF=docRef.textFrames.add();
    newTF.contents=text;
    newTF.top=y;
    newTF.left=x;
    newTF.textRange.characterAttributes.size = size;
  switch (forw){
   case 0:
  newTF.textRange.paragraphAttributes.justification = Justification.LEFT;
  break;
   case 1:
  newTF.textRange.paragraphAttributes.justification = Justification.CENTER;
  break;   
   case 2:
  newTF.textRange.paragraphAttributes.justification = Justification.RIGHT;
  break;   
  
  }
}

function readXML(){

dataBase=File('z:/zakaz.xml');  //место хранения xml-файла с данными
dataBase.open('read');

      var i;
      var x;
      var dlinaStr;
      
      for(i=0;i<=30;i++){
      var str = dataBase.readln();

      if(str.indexOf('shirina') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        shirina=str.substr(9,x);   
                        }
      if(str.indexOf('vysota') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-18;
        vysota=str.substr(8,x);   
//        alert (vysota);
                        }
      if(str.indexOf('manager') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        manager=str.substr(9,x);   
    switch (manager){
     case ('...'): //имя менеджера в БД
      manager="..."; // Подстановка реальных ФИО менеджера
      break;
     } 
                        }
      if(str.indexOf('zakazchik') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-23;
        zakazchik=str.substr(11,x);   
                        }
      if(str.indexOf('nazvanie') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-21;
        nazvanie=str.substr(10,x);   
                        }           
      if(str.indexOf('dno') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-11;
        dno=parseInt(str.substr(5,x));   
                        }           
      if(str.indexOf('verh') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-13;
        verh=parseInt(str.substr(6,x));   
                        }
      if(str.indexOf('bokovye') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        bokovye=parseInt(str.substr(9,x));   
                        }
      if(str.indexOf('nozh') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-13;
        nozh=parseInt(str.substr(6,x));   
        if (nozh>99){
        nozh=nozh/10;
        }
                        }
      if(str.indexOf('printer') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        printer=str.substr(9,x);   
                        } 
    if(str.indexOf('namotka') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        namotk=str.substr(9,x);   
                        }
      if(str.indexOf('formula') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        formula=str.substr(9,x);   
                        }    
      if(str.indexOf('izdelie') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-19;
        izdelie=str.substr(9,x);   
                        }   
      if(str.indexOf('raport') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-17;
        raport=str.substr(8,x);   
                        } 
      if(str.indexOf('material') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-21;
        material=str.substr(10,x);   
                    }             

      if(str.indexOf('colors') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-17;
        colors=str.substr(8,x)

//        alert (colors);
        colors=colors.replace(/ P.:/g,',');
//        alert (colors);
        colors=colors.replace(/Pantone/g,'');
//        alert (colors);
        colors=colors.replace(/pantone/g,'');
//        alert (colors);
        colors=colors.replace(new RegExp(/\s/g),'');
        allColors=colors.split(new RegExp(/\+|\,/));   
//        alert (allColors);

        if (allColors[allColors.length-1]==''){
        allColors.pop();
}
else {
}

                }

        if(str.indexOf('technomer') != -1) {
        dlinaStr=str.length;
        x=dlinaStr-23;
        technomer=str.substr(11,x);   
                        }
                        
                        }
}

function chngArt(num, heig, widt){
cur=docRef.artboards.length-1;
if (num<=cur){
 artBoar=docRef.artboards[num];
 artX=artBoar.artboardRect[0];
 artY=artBoar.artboardRect[1];
 artShirina=artX*sizeIndex+heig/sizeIndex;
 artWysota=artY*sizeIndex-widt/sizeIndex;
 artSize=[artX,artY,artShirina,artWysota];
 docRef.artboards[num].artboardRect=artSize;
 }
else {
    var artX=(docRef.artboards[num-1].artboardRect[2]-docRef.artboards[0].artboardRect[0])*sizeIndex+shift;
    var artY=(docRef.artboards[num-1].artboardRect[1])*sizeIndex;
    var artH=artX+heig;
    var artW=artY-widt;
    artX=artX/sizeIndex;
    artY=artY/sizeIndex;
    artH=artH/sizeIndex;
    artW=artW/sizeIndex;
    artSize=[artX,artY,artH,artW];
    docRef.artboards.add(artSize);
}

}


function addColor(ColorName){
if (ColorName=="Праймер"){
      spotL=100;
      spotA=30;
      spotB=30;

var Newcolor = new LabColor();
Newcolor.l = spotL;
Newcolor.a = spotA;
Newcolor.b = spotB;
Newcolor.typename=ColorName;
var swatch = docRef.spots.add();
swatch.colorType=ColorModel.SPOT;
swatch.color = Newcolor;
swatch.name = "Primer";
}

if (ColorName=="white"){
      spotL=100;
      spotA=-40;
      spotB=0;

var Newcolor = new LabColor();
Newcolor.l = spotL;
Newcolor.a = spotA;
Newcolor.b = spotB;
Newcolor.typename=ColorName;
var swatch = docRef.spots.add();
swatch.colorType=ColorModel.SPOT;
swatch.color = Newcolor;
swatch.name = ColorName;
}
else {

var pathJSX=$.fileName;
pathJSX = pathJSX.substring(0,pathJSX.lastIndexOf("\\"));
pathJSX = replace_string(pathJSX,'\\','/');
pathJSX = '/'+replace_string(pathJSX,':','');


dataBase=File(pathJSX+'/pantone.xml');
  dataBase.open('read');
      var i;
      var x;
      var dlinaStr;
      
      for(i=0;i<=dataBase.length;i++){
      var str = dataBase.readln();
    if (ColorName=="072"){
      ColorName="Blue 072";
    }
    if (ColorName=="021"){
      ColorName="Orange 021";
    }
    if (ColorName=="1"){
      ColorName="Cool Gray 1";
    }
      if(str.indexOf('<td>'+ColorName+'</td>') != -1) {

      startchar=str.indexOf('<td>'+ColorName+'</td>');
      tempcolor=str.substr(startchar,200);
      lastchar=tempcolor.indexOf('</tr>');
      tempcolor=str.substr(startchar,lastchar);
      fullcolor=tempcolor.replace(/<td>/g,"");
      colorz=fullcolor.split('</td>');
      spotL=colorz[2];
      spotA=colorz[3];
      spotB=colorz[4];
// Create the new color for the swatch
var Newcolor = new LabColor();
Newcolor.l = spotL;
Newcolor.a = spotA;
Newcolor.b = spotB;
Newcolor.typename=ColorName;

// Create the new swatch using the above color
var swatch = docRef.spots.add();
swatch.colorType=ColorModel.SPOT;
swatch.color = Newcolor;
swatch.name = "PANTONE "+ColorName+" C";


      break;
                        }

                        }
}
}
function replace_string(txt,cut_str,paste_str){ 
var f=0;
var ht='';
ht = ht + txt;
f=ht.indexOf(cut_str);
while (f!=-1){ 
f=ht.indexOf(cut_str);
if (f>0){
ht = ht.substr(0,f) + paste_str + ht.substr(f+cut_str.length);
};
};
return ht;
};