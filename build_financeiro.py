/**
 * Let's Go — Agência de Viagens · v3.0
 * Correções: IDs concorrentes, validações, atualização de clientes, histórico real,
 * operações atómicas, geração segura de PDF e mensagens de erro visíveis.
 */
var SS = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Let's Go")
    .addItem("💾 Guardar reserva", "guardarReserva")
    .addItem("👤 Atualizar cliente", "atualizarCliente")
    .addItem("➕ Adicionar serviço", "adicionarServico")
    .addItem("➕ Adicionar passageiro", "adicionarPassageiro")
    .addItem("💳 Registar pagamento", "registarPagamento")
    .addItem("✅ Aprovar reserva / definir venda", "aprovarReserva")
    .addSeparator().addItem("📄 Gerar PDF do comprovativo", "gerarPDF").addToUi();
}

function _sh(n){var s=SS.getSheetByName(n);if(!s)throw new Error("Aba em falta: "+n);return s;}
function _v(a){return _sh("Formulário").getRange(a).getValue();}
function _txt(v){return String(v==null?"":v).trim();}
function _money(v){return typeof v==="number"&&isFinite(v)?v:NaN;}
function _date(v){return v instanceof Date&&!isNaN(v.getTime());}
function _user(){try{return Session.getActiveUser().getEmail()||"Utilizador";}catch(e){return "Utilizador";}}
function _findRow(sheet,id,col){if(sheet.getLastRow()<2)return -1;var a=sheet.getRange(2,col,sheet.getLastRow()-1,1).getValues();for(var i=0;i<a.length;i++)if(String(a[i][0])===String(id))return i+2;return -1;}
function _alert(t,m){SpreadsheetApp.getUi().alert(t,m,SpreadsheetApp.getUi().ButtonSet.OK);}

function proximoId(sheet,col,prefixo){
  var max=0,last=sheet.getLastRow();
  if(last>=2)sheet.getRange(2,col,last-1,1).getDisplayValues().forEach(function(x){var m=String(x[0]||"").match(/(\d+)\s*$/);if(m)max=Math.max(max,Number(m[1]));});
  return prefixo+String(max+1).padStart(4,"0");
}

function _historico(acao,aba,referencia,antes,depois){
  var h=_sh("Histórico"),r=Math.max(h.getLastRow()+1,4),agora=new Date();
  h.getRange(r,1,1,8).setValues([[agora,agora,_user(),acao,aba,referencia,_txt(antes),_txt(depois)]]);
}

function _fPassageiro(s,r){
  s.getRange(r,8).setFormula('=IFERROR(IF(OR($G'+r+'="",$A'+r+'=""),"",DATEDIF($G'+r+',INDEX(Reservas!$I:$I,MATCH($A'+r+',Reservas!$A:$A,0)),"Y")),"")');
  s.getRange(r,10).setFormula('=IF(OR($E'+r+'="",$A'+r+'=""),"",IFERROR(IF(INDEX(Reservas!$J:$J,MATCH($A'+r+',Reservas!$A:$A,0))="","",IF($E'+r+'<INDEX(Reservas!$J:$J,MATCH($A'+r+',Reservas!$A:$A,0)),"🔴 Expira antes do regresso",IF($E'+r+'<EDATE(INDEX(Reservas!$J:$J,MATCH($A'+r+',Reservas!$A:$A,0)),Configurações!$L$4),"🟡 Validade curta","🟢 OK"))),""))');
}

function guardarReserva(){
  var lock=LockService.getDocumentLock();lock.waitLock(30000);
  var reservaCriada=null,passCriado=null,clienteCriado=null;
  try{
    var form=_sh("Formulário"),clientes=_sh("Clientes"),reservas=_sh("Reservas"),pass=_sh("Passageiros");
    var idCliente=_txt(_v("C6")),pacote=_txt(_v("C16")),ini=_v("C17"),fim=_v("C18"),pax=Number(_v("C19")),tipo=_txt(_v("C21")),prest=Number(_v("C22"));
    if(!pacote||!_date(ini)||!_date(fim)||fim<ini||!Number.isInteger(pax)||pax<1)throw new Error("Preenche pacote, datas válidas e número de passageiros maior que zero.");
    if(!tipo)throw new Error("Escolhe o tipo de pagamento.");
    if(!Number.isInteger(prest)||prest<1)throw new Error("O número de prestações deve ser pelo menos 1.");
    if(tipo==="A pronto"&&prest!==1)throw new Error("Pagamento a pronto deve ter exatamente 1 prestação.");
    var tit={};
    if(!idCliente){
      var nome=_txt(_v("C7"));if(!nome)throw new Error("Escolhe um cliente ou escreve o nome do novo cliente.");
      idCliente=proximoId(clientes,1,"C-");var rc=clientes.getLastRow()+1;clienteCriado=rc;
      clientes.getRange(rc,1,1,11).setValues([[idCliente,nome,_v("C8"),_v("C9"),_v("C10"),_v("C11"),_v("C12"),_v("C13"),_v("C14"),new Date(),""]]);
      clientes.getRange(rc,12).setFormula('=IF($I'+rc+'="","",DATEDIF($I'+rc+',TODAY(),"Y"))');tit={nome:nome,doc:_v("C11"),val:_v("C12"),nac:_v("C13"),nasc:_v("C14")};
    }else{
      var cr=_findRow(clientes,idCliente,1);if(cr<0)throw new Error("O cliente selecionado já não existe.");
      var c=clientes.getRange(cr,1,1,9).getValues()[0];tit={nome:c[1],doc:c[5],val:c[6],nac:c[7],nasc:c[8]};
    }
    if(!_txt(tit.nome))throw new Error("O titular não tem nome válido.");
    var idReserva=proximoId(reservas,1,"R-"+new Date().getFullYear()+"-"),r=reservas.getLastRow()+1;reservaCriada=r;
    reservas.getRange(r,1,1,32).setValues([[idReserva,new Date(),idCliente,"",pacote,"","","",ini,fim,pax,_v("C20")||"Em orçamento","Aguarda aprovação","","","","",tipo,prest,"","","","","","",_v("C23"),_user(),"","","","",_v("C24")]]);
    var fs={4:'=IFERROR(INDEX(Clientes!$B:$B,MATCH($C'+r+',Clientes!$A:$A,0)),"")',6:'=IFERROR(INDEX(Pacotes!$B:$B,MATCH($E'+r+',Pacotes!$A:$A,0)),"")',7:'=IFERROR(INDEX(Pacotes!$C:$C,MATCH($E'+r+',Pacotes!$A:$A,0)),"")',8:'=IFERROR(INDEX(Pacotes!$D:$D,MATCH($E'+r+',Pacotes!$A:$A,0)),"")',14:'=IF($A'+r+'="","",SUMIF(Serviços!$A:$A,$A'+r+',Serviços!$N:$N))',16:'=IF(OR($O'+r+'="",$N'+r+'=""),"",$O'+r+'-$N'+r+')',20:'=IF($A'+r+'="","",COUNTIF(Registos!$A:$A,$A'+r+'))',21:'=IF(OR($O'+r+'="",$S'+r+'="",$S'+r+'=0),"",ROUND($O'+r+'/$S'+r+',2))',22:'=IF($A'+r+'="","",SUMIF(Registos!$A:$A,$A'+r+',Registos!$C:$C))',23:'=IF($O'+r+'="","",$O'+r+'-$V'+r+')',24:'=IF($S'+r+'="","",MAX(0,$S'+r+'-IF($T'+r+'="",0,$T'+r+')))',25:'=IF(OR($O'+r+'="",$O'+r+'=0),"",MIN(1,$V'+r+'/$O'+r+'))'};
    Object.keys(fs).forEach(function(k){reservas.getRange(r,Number(k)).setFormula(fs[k]);});
    var pr=pass.getLastRow()+1;passCriado=pr;pass.getRange(pr,1,1,9).setValues([[idReserva,tit.nome,"Sim",tit.doc,tit.val,tit.nac,tit.nasc,"",""]]);_fPassageiro(pass,pr);
    _historico("CRIAR RESERVA","Reservas",idReserva,"",JSON.stringify({cliente:idCliente,pacote:pacote,pax:pax}));
    form.getRangeList(["C6:C14","C16:C24"]).clearContent();SS.toast("Reserva "+idReserva+" guardada.","Let's Go",5);
  }catch(e){
    if(passCriado)_sh("Passageiros").deleteRow(passCriado);if(reservaCriada)_sh("Reservas").deleteRow(reservaCriada);if(clienteCriado)_sh("Clientes").deleteRow(clienteCriado);
    _alert("Não foi possível guardar",e.message);
  }finally{lock.releaseLock();}
}

function atualizarCliente(){
  try{var s=_sh("Clientes"),id=_txt(_v("C6")),r=_findRow(s,id,1);if(!id||r<0)throw new Error("Escolhe um cliente existente em C6.");var antes=s.getRange(r,2,1,8).getDisplayValues()[0].join(" | ");
    var vals=[_v("C7"),_v("C8"),_v("C9"),_v("C10"),_v("C11"),_v("C12"),_v("C13"),_v("C14")];var atual=s.getRange(r,2,1,8).getValues()[0];for(var i=0;i<vals.length;i++)if(vals[i]!=="")atual[i]=vals[i];s.getRange(r,2,1,8).setValues([atual]);
    _historico("ATUALIZAR CLIENTE","Clientes",id,antes,s.getRange(r,2,1,8).getDisplayValues()[0].join(" | "));SS.toast("Cliente atualizado.","Let's Go",4);
  }catch(e){_alert("Não foi possível atualizar",e.message);}
}

function adicionarServico(){
  try{var s=_sh("Serviços"),idr=_txt(_v("C30")),tipo=_txt(_v("C31")),custo=_money(_v("C42"));if(_findRow(_sh("Reservas"),idr,1)<0)throw new Error("Escolhe uma reserva válida.");if(!tipo)throw new Error("Escolhe o tipo de serviço.");if(!isFinite(custo)||custo<0)throw new Error("Indica um custo válido, igual ou superior a zero.");
    var r=s.getLastRow()+1;s.getRange(r,1,1,14).setValues([[idr,tipo,_v("C32"),_v("C33"),"",_v("C34"),_v("C35"),_v("C36"),_v("C37"),_v("C38"),_v("C39"),_v("C40"),_v("C41"),custo]]);s.getRange(r,5).setFormula('=IFERROR(INDEX(Fornecedores!$B:$B,MATCH($D'+r+',Fornecedores!$A:$A,0)),"")');_historico("ADICIONAR SERVIÇO","Serviços",idr,"",tipo+" | "+custo);_sh("Formulário").getRange("C30:C42").clearContent();
  }catch(e){_alert("Não foi possível adicionar",e.message);}
}

function adicionarPassageiro(){
  try{var s=_sh("Passageiros"),idr=_txt(_v("C48")),nome=_txt(_v("C49"));if(_findRow(_sh("Reservas"),idr,1)<0||!nome)throw new Error("Escolhe uma reserva válida e indica o nome.");var r=s.getLastRow()+1;s.getRange(r,1,1,9).setValues([[idr,nome,"Não",_v("C50"),_v("C51"),_v("C52"),_v("C53"),"",_v("C54")]]);_fPassageiro(s,r);_historico("ADICIONAR PASSAGEIRO","Passageiros",idr,"",nome);_sh("Formulário").getRange("C48:C54").clearContent();}catch(e){_alert("Não foi possível adicionar",e.message);}
}

function registarPagamento(){
  var lock=LockService.getDocumentLock();lock.waitLock(30000);try{var idr=_txt(_v("C60")),valor=_money(_v("C62")),res=_sh("Reservas"),rr=_findRow(res,idr,1);if(rr<0)throw new Error("Escolhe uma reserva válida.");if(!isFinite(valor)||valor<=0)throw new Error("O pagamento deve ser superior a zero.");if(res.getRange(rr,12).getValue()==="Cancelada")throw new Error("Não é possível pagar uma reserva cancelada.");SpreadsheetApp.flush();var pend=Number(res.getRange(rr,23).getValue());if(isFinite(pend)&&valor>pend+0.005)throw new Error("O pagamento é superior ao valor pendente ("+pend+" €).");
    var s=_sh("Registos"),r=s.getLastRow()+1;s.getRange(r,1,1,5).setValues([[idr,_v("C61")||new Date(),valor,_v("C63"),_v("C64")]]);_historico("REGISTAR PAGAMENTO","Registos",idr,"",valor+" €");_sh("Formulário").getRange("C60:C64").clearContent();SS.toast("Pagamento registado.","Let's Go",5);
  }catch(e){_alert("Não foi possível registar",e.message);}finally{lock.releaseLock();}
}

function aprovarReserva(){
  try{var idr=_txt(_v("C70")),venda=_money(_v("C71")),s=_sh("Reservas"),r=_findRow(s,idr,1);if(r<0)throw new Error("Reserva não encontrada.");SpreadsheetApp.flush();var custo=Number(s.getRange(r,14).getValue());if(!isFinite(venda)||venda<=0)throw new Error("O valor de venda deve ser superior a zero.");if(isFinite(custo)&&venda<custo){var b=SpreadsheetApp.getUi().alert("Margem negativa","A venda é inferior ao custo. Aprovar mesmo assim?",SpreadsheetApp.getUi().ButtonSet.YES_NO);if(b!==SpreadsheetApp.getUi().Button.YES)return;}
    var antes=s.getRange(r,13,1,5).getDisplayValues()[0].join(" | ");s.getRange(r,15).setValue(venda);s.getRange(r,13).setValue("Aprovado");s.getRange(r,17).setValue(_v("C72")||_user());if(s.getRange(r,12).getValue()==="Em orçamento")s.getRange(r,12).setValue("Confirmada");_historico("APROVAR RESERVA","Reservas",idr,antes,"Venda: "+venda);_sh("Formulário").getRange("C70:C72").clearContent();
  }catch(e){_alert("Não foi possível aprovar",e.message);}
}

function gerarPDF(){
  var dup=null,temp=null;try{var fid=_sh("Ficha de Reserva").getRange("C9").getValue();if(!fid)throw new Error("Escolhe primeiro uma reserva na Ficha de Reserva.");var src=_sh("Comprovativo");dup=src.copyTo(SS);SpreadsheetApp.flush();var rng=dup.getDataRange();rng.copyTo(rng,{contentsOnly:true});temp=SpreadsheetApp.create("TEMP_Comprovativo_"+fid);dup.copyTo(temp).setName("Comprovativo");temp.getSheets().forEach(function(s){if(s.getName()!=="Comprovativo")temp.deleteSheet(s);});SpreadsheetApp.flush();var folder=_obterPastaPDF();var pdf=DriveApp.getFileById(temp.getId()).getAs(MimeType.PDF).setName("Comprovativo_"+fid+".pdf");var file=folder.createFile(pdf);_historico("GERAR PDF","Comprovativo",fid,"",file.getName());SpreadsheetApp.getUi().alert("PDF criado:\n\n"+file.getUrl());
  }catch(e){_alert("Não foi possível gerar o PDF",e.message);}finally{try{if(temp)DriveApp.getFileById(temp.getId()).setTrashed(true);}catch(e){}try{if(dup)SS.deleteSheet(dup);}catch(e){}}
}
function _obterPastaPDF(){var it=DriveApp.getFoldersByName("Let's Go - Comprovativos");return it.hasNext()?it.next():DriveApp.createFolder("Let's Go - Comprovativos");}

function onEdit(e){
  try{if(!e||!e.range)return;var n=e.range.getSheet().getName();if(["Reservas","Serviços","Passageiros","Clientes","Registos"].indexOf(n)<0)return;var novo=e.range.getDisplayValue();_historico("EDIÇÃO MANUAL",n,e.range.getA1Notation(),e.oldValue||"",novo);}catch(err){console.error(err);}
}
