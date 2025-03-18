function copyColumnsTest(){
    let ins = new CopyColumns(SH_ID_ORIG,SH_ID_DEST);
  
    ins.targetColumns = 'a';
    ins.targetColumns = 'c';
    ins.targetColumns = 'e';

    ins.copyColumns();
  
}