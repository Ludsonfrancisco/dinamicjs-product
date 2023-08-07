let selectedFile

document.getElementById('input').addEventListener('change', (event) => {
     selectedFile = event.target.files[0]
})



document.getElementById('button').addEventListener('click', () => {
     
     if (selectedFile) {
          let fileReader = new FileReader()
          fileReader.readAsBinaryString(selectedFile)
          fileReader.onload = (event) => {
               let data = event.target.result
               let workbook = XLSX.read(data, { type: 'binary' })
               // console.log(workbook)
               workbook.SheetNames.forEach(sheet => {
                    let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet])


                    //===================== STATUS DA ORDEM =============

                    const statusEmRota = item => item.Status === 'em rota'
                    const statusConcluida = item => item.Status === 'Concluída'
                    const statusIniciada = item => item.Status === 'Iniciada'
                    const statusNaoIniciada = item => item.Status === 'Não Iniciada'
                    const statusNaoConcluida = item => item.Status === 'Não Concluída'



                    //====================== SLOT ORDEM ====================
                    const slot10 = item => item['Janela de Serviço'] === '08:30 - 10:30'
                    const slot12 = item => item['Janela de Serviço'] === '10 - 12'
                    const slot15 = item => item['Janela de Serviço'] === '13:00 - 15:30'
                    const slot18 = item => item['Janela de Serviço'] === '15:30 - 18:00'

                    

                    //=======================FILTROS 



                    const conIniNin = item => {
                         if (item.Status === 'Concluída' || item.Status === 'Iniciada' || item.Status === 'Não Iniciada')
                              return item
                    }

                    const metalico = (item) => {
                         if (item['Habilidades de Trabalho'] === 'Reparo Linha(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Banda(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo TV(1/100)')

                              return item
                    }

                    const gpon = (item) => {
                         if (item['Habilidades de Trabalho'] === 'Reparo Banda FB Alto Valor(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Banda FB(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Linha FB Alto Valor(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo Linha FB(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo TV FB(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo TV FB Alto Valor(1/100)' ||
                              item['Habilidades de Trabalho'] === 'Reparo FTTA(1/100)' ||     
                              item['Habilidades de Trabalho'] === '' && 
                                                       ( item["Tipo de Atividade"] === 'Defeito Banda Larga' ||
                                                       item["Tipo de Atividade"] === 'Defeito TV' ||
                                                       item["Tipo de Atividade"] === 'Defeito Linha' 
                                                       )   
                              ) 
                              

                              return item
                    }


                    
               

                    const prev = (item => {
                         if ((item['Detalhe da Atividade'].charAt(0) === 'P' || item['Detalhe da Atividade'].charAt(0) === 'p') &&
                              (item['Detalhe da Atividade'].charAt(1) === 'R' || item['Detalhe da Atividade'].charAt(1) === 'r') &&
                              (item['Detalhe da Atividade'].charAt(2) === 'E' || item['Detalhe da Atividade'].charAt(2) === 'e') &&
                              (item['Detalhe da Atividade'].charAt(3) === 'V' || item['Detalhe da Atividade'].charAt(3) === 'v')
                         )
                              return item
                    })



                    // ANTES DE RETIRAR PREV

                    const data = rowObject

                    // FUNÇAO PARA RETIRAR PREVENTIVA =================

                    function removeItem(arr, prop, value) {
                         prop.toUpperCase()
                         return arr.filter(function (i) { return i[prop] !== value })
                    }


                    rowObject = removeItem(rowObject, 'Detalhe da Atividade', 'PREV/BANDA LARGA')
                    rowObject = removeItem(rowObject, 'Detalhe da Atividade', 'PREV/PÓS CONTATO')
                    rowObject = removeItem(rowObject, 'Detalhe da Atividade', 'Prev/Pós Contato')


                    //===================== DADOS CIDADES                   
                    let dataArc = rowObject.filter(item => item.Cidade === 'ARACRUZ')
                    let dataCim = rowObject.filter(item => item.Cidade === 'CACHOEIRO DE ITAPEMIRIM')
                    let dataCca = rowObject.filter(item => item.Cidade === 'CARIACICA')
                    let dataCna = rowObject.filter(item => item.Cidade === 'COLATINA')
                    let dataGri = rowObject.filter(item => item.Cidade === 'GUARAPARI')
                    let dataLns = rowObject.filter(item => item.Cidade === 'LINHARES')
                    let dataSmj = rowObject.filter(item => item.Cidade === 'SANTA MARIA DE JETIBA')
                    let dataSmt = rowObject.filter(item => item.Cidade === 'SAO MATEUS')
                    let dataSea = rowObject.filter(item => item.Cidade === 'SERRA')
                    let dataVva = rowObject.filter(item => item.Cidade === 'VILA VELHA')
                    let dataVta = rowObject.filter(item => item.Cidade === 'VITORIA')
                    let dataVia = rowObject.filter(item => item.Cidade === 'VIANA')



                    // ============= DADOS POR CIDADE GPON 

                    const dataGpon = rowObject.filter(gpon)
                    const gponArc = dataArc.filter(gpon)
                    const gponCim = dataCim.filter(gpon)
                    const gponCca = dataCca.filter(gpon)
                    const gponCna = dataCna.filter(gpon)
                    const gponGri = dataGri.filter(gpon)
                    const gponLns = dataLns.filter(gpon)
                    const gponSmj = dataSmj.filter(gpon)
                    const gponSmt = dataSmt.filter(gpon)
                    const gponSea = dataSea.filter(gpon)
                    const gponVva = dataVva.filter(gpon)
                    const gponVta = dataVta.filter(gpon)
                    const gponVia = dataVia.filter(gpon)


                    //console.log(gponAtual)

                    


               // ====CREATE CABEÇALHO TABELA PRODUÇAO GPON ====
                    // const titleProducao = document.createElement('h1')
                    // titleProducao.className = 'title-producao'
                    // titleProducao.innerHTML = 'PRODUÇÃO'
                    // const tProducao = document.getElementById("producao")
                    // tProducao.appendChild(titleProducao)


                    const titleGpon = document.createElement('span')
                    titleGpon.innerHTML = 'GPON'
                    const tGpon = document.getElementById("gpon-titulo")
                    tGpon.append(titleGpon)




                    const tdCidade = document.createElement('td')
                    tdCidade.className = 'tdCidade'
                    tdCidade.innerHTML = 'CIDADE'    
                    tdCidade.colSpan = '1'                

                    const tdIniciada = document.createElement('td')
                    tdIniciada.className = 'tdIniciada'
                    tdIniciada.innerHTML = 'INICIADA'
                    tdIniciada.colSpan = '4'

                    const tdNaoiniciada = document.createElement('td')
                    tdNaoiniciada.className = 'tdNin'
                    tdNaoiniciada.innerHTML = 'NÃO INICIADA'
                    tdNaoiniciada.colSpan = '4'

               

                    const tabela = document.getElementById('cabecalho')
                    tabela.append(tdCidade)
                    tabela.append(tdIniciada)
                    tabela.append(tdNaoiniciada)






                    // TABELA DE HORARIO DE SLOT 
                    const tdSlot = document.createElement('td')
                    tdSlot.className = 'tdSlot'
                    tdSlot.innerHTML = 'SLOT'
                    const tab = document.getElementById('slot')
                    tab.append(tdSlot)

                    const tdIni10 = document.createElement('td')
                    tdIni10.innerHTML = '10:30h'
                    tdIni10.className = 'tdSlot'
                    const colIni10 = document.getElementById('slot')
                    colIni10.append(tdIni10)

                    const tdIni12 = document.createElement('td')
                    tdIni12.innerHTML = '12:00h'
                    tdIni12.className = 'tdSlot'
                    const colIni12 = document.getElementById('slot')
                    colIni12.append(tdIni12)

                    const tdIni15 = document.createElement('td')
                    tdIni15.innerHTML = '15:30h'
                    tdIni15.className = 'tdSlot'
                    const colIni15 = document.getElementById('slot')
                    colIni15.append(tdIni15)

                    const tdIni18 = document.createElement('td')
                    tdIni18.innerHTML = '18:00h'
                    tdIni18.className = 'tdSlot'
                    const colIni18 = document.getElementById('slot')
                    colIni18.append(tdIni18)



                    const tdNin10 = document.createElement('td')
                    tdNin10.innerHTML = '10:30h'
                    tdNin10.className = 'tdSlot'
                    const colNin10 = document.getElementById('slot')
                    colNin10.append(tdNin10)

                    const tdNin12 = document.createElement('td')
                    tdNin12.innerHTML = '12:00h'
                    tdNin12.className = 'tdSlot'
                    const colNin12 = document.getElementById('slot')
                    colNin12.append(tdNin12)

                    const tdNin15 = document.createElement('td')
                    tdNin15.innerHTML = '15:30h'
                    tdNin15.className = 'tdSlot'
                    const colNin15 = document.getElementById('slot')
                    colNin15.append(tdNin15)

                    const tdNin18 = document.createElement('td')
                    tdNin18.innerHTML = '18:00h'
                    tdNin18.className = 'tdSlot'
                    const colNin18 = document.getElementById('slot')
                    colNin18.append(tdNin18)

               

                    // ======== CREATE CIDADE GPON ==


                    // ARACRUZ
                    const tdArcGpon = document.createElement('td')
                    tdArcGpon.innerHTML = 'ARACRUZ'
                    const colArcGpon = document.getElementById('arc')
                    colArcGpon.append(tdArcGpon)


                    // iniado
                    const tdIniArcSlot10 = document.createElement('td')
                    tdIniArcSlot10.innerHTML = gponArc.filter(slot10).filter(statusIniciada).length
                    const conIniArcSlot10 = document.getElementById('arc')
                    conIniArcSlot10.append(tdIniArcSlot10)

                    const tdIniArcSlot12 = document.createElement('td')
                    tdIniArcSlot12.innerHTML = gponArc.filter(slot12).filter(statusIniciada).length
                    const conIniArcSlot12 = document.getElementById('arc')
                    conIniArcSlot12.append(tdIniArcSlot12)

                    const tdIniArcSlot15 = document.createElement('td')
                    tdIniArcSlot15.innerHTML = gponArc.filter(slot15).filter(statusIniciada).length
                    const conIniArcSlot15 = document.getElementById('arc')
                    conIniArcSlot15.append(tdIniArcSlot15)

                    const tdIniArcSlot18 = document.createElement('td')
                    tdIniArcSlot18.innerHTML = gponArc.filter(slot18).filter(statusIniciada).length
                    const conIniArcSlot18 = document.getElementById('arc')
                    conIniArcSlot18.append(tdIniArcSlot18)


                    // não iniciada
                    
                    const tdNinArcSlot10 = document.createElement('td')
                    tdNinArcSlot10.innerHTML = gponArc.filter(slot10).filter(statusNaoIniciada).length
                    const conNinArcSlot10 = document.getElementById('arc')
                    conNinArcSlot10.append(tdNinArcSlot10)

                    const tdNinArcSlot12 = document.createElement('td')
                    tdNinArcSlot12.innerHTML = gponArc.filter(slot12).filter(statusNaoIniciada).length
                    const conNinArcSlot12 = document.getElementById('arc')
                    conNinArcSlot12.append(tdNinArcSlot12)

                    const tdNinArcSlot15 = document.createElement('td')
                    tdNinArcSlot15.innerHTML = gponArc.filter(slot15).filter(statusNaoIniciada).length
                    const conNinArcSlot15 = document.getElementById('arc')
                    conNinArcSlot15.append(tdNinArcSlot15)

                    const tdNinArcSlot18 = document.createElement('td')
                    tdNinArcSlot18.innerHTML = gponArc.filter(slot18).filter(statusNaoIniciada).length
                    const conNinArcSlot18 = document.getElementById('arc')
                    conNinArcSlot18.append(tdNinArcSlot18)



                    // CACHOEIRO
                    const tdCimGpon = document.createElement('td')
                    tdCimGpon.innerHTML = 'CACHOEIRO'
                    const colCimGpon = document.getElementById('cim')
                    colCimGpon.append(tdCimGpon)


                    // iniado
                    const tdIniCimSlot10 = document.createElement('td')
                    tdIniCimSlot10.innerHTML = gponCim.filter(slot10).filter(statusIniciada).length
                    const conIniCimSlot10 = document.getElementById('cim')
                    conIniCimSlot10.append(tdIniCimSlot10)

                    const tdIniCimSlot12 = document.createElement('td')
                    tdIniCimSlot12.innerHTML = gponCim.filter(slot12).filter(statusIniciada).length
                    const conIniCimSlot12 = document.getElementById('cim')
                    conIniCimSlot12.append(tdIniCimSlot12)

                    const tdIniCimSlot15 = document.createElement('td')
                    tdIniCimSlot15.innerHTML = gponCim.filter(slot15).filter(statusIniciada).length
                    const conIniCimSlot15 = document.getElementById('cim')
                    conIniCimSlot15.append(tdIniCimSlot15)

                    const tdIniCimSlot18 = document.createElement('td')
                    tdIniCimSlot18.innerHTML = gponCim.filter(slot18).filter(statusIniciada).length
                    const conIniCimSlot18 = document.getElementById('cim')
                    conIniCimSlot18.append(tdIniCimSlot18)
               

                    // não iniciada
                    
                    const tdNinCimSlot10 = document.createElement('td')
                    tdNinCimSlot10.innerHTML = gponCim.filter(slot10).filter(statusNaoIniciada).length
                    const conNinCimSlot10 = document.getElementById('cim')
                    conNinCimSlot10.append(tdNinCimSlot10)

                    const tdNinCimSlot12 = document.createElement('td')
                    tdNinCimSlot12.innerHTML = gponCim.filter(slot12).filter(statusNaoIniciada).length
                    const conNinCimSlot12 = document.getElementById('cim')
                    conNinCimSlot12.append(tdNinCimSlot12)

                    const tdNinCimSlot15 = document.createElement('td')
                    tdNinCimSlot15.innerHTML = gponCim.filter(slot15).filter(statusNaoIniciada).length
                    const conNinCimSlot15 = document.getElementById('cim')
                    conNinCimSlot15.append(tdNinCimSlot15)

                    const tdNinCimSlot18 = document.createElement('td')
                    tdNinCimSlot18.innerHTML = gponCim.filter(slot18).filter(statusNaoIniciada).length
                    const conNinCimSlot18 = document.getElementById('cim')
                    conNinCimSlot18.append(tdNinCimSlot18)


                    // CARIACICA
                    const tdCcaGpon = document.createElement('td')
                    tdCcaGpon.innerHTML = 'CARIACICA'
                    const colCcaGpon = document.getElementById('cca')
                    colCcaGpon.append(tdCcaGpon)


                        // iniado
                    const tdIniCcaSlot10 = document.createElement('td')
                    tdIniCcaSlot10.innerHTML = gponCca.filter(slot10).filter(statusIniciada).length
                    const conIniCcaSlot10 = document.getElementById('cca')
                    conIniCcaSlot10.append(tdIniCcaSlot10)

                    const tdIniCcaSlot12 = document.createElement('td')
                    tdIniCcaSlot12.innerHTML = gponCca.filter(slot12).filter(statusIniciada).length
                    const conIniCcaSlot12 = document.getElementById('cca')
                    conIniCcaSlot12.append(tdIniCcaSlot12)

                    const tdIniCcaSlot15 = document.createElement('td')
                    tdIniCcaSlot15.innerHTML = gponCca.filter(slot15).filter(statusIniciada).length
                    const conIniCcaSlot15 = document.getElementById('cca')
                    conIniCcaSlot15.append(tdIniCcaSlot15)

                    const tdIniCcaSlot18 = document.createElement('td')
                    tdIniCcaSlot18.innerHTML = gponCca.filter(slot18).filter(statusIniciada).length
                    const conIniCcaSlot18 = document.getElementById('cca')
                    conIniCcaSlot18.append(tdIniCcaSlot18)
               

                    // não iniciada
                    
                    const tdNinCcaSlot10 = document.createElement('td')
                    tdNinCcaSlot10.innerHTML = gponCca.filter(slot10).filter(statusNaoIniciada).length
                    const conNinCcaSlot10 = document.getElementById('cca')
                    conNinCcaSlot10.append(tdNinCcaSlot10)

                    const tdNinCcaSlot12 = document.createElement('td')
                    tdNinCcaSlot12.innerHTML = gponCca.filter(slot12).filter(statusNaoIniciada).length
                    const conNinCcaSlot12 = document.getElementById('cca')
                    conNinCcaSlot12.append(tdNinCcaSlot12)

                    const tdNinCcaSlot15 = document.createElement('td')
                    tdNinCcaSlot15.innerHTML = gponCca.filter(slot15).filter(statusNaoIniciada).length
                    const conNinCcaSlot15 = document.getElementById('cca')
                    conNinCcaSlot15.append(tdNinCcaSlot15)

                    const tdNinCcaSlot18 = document.createElement('td')
                    tdNinCcaSlot18.innerHTML = gponCca.filter(slot18).filter(statusNaoIniciada).length
                    const conNinCcaSlot18 = document.getElementById('cca')
                    conNinCcaSlot18.append(tdNinCcaSlot18)

                    

                    // COLATINA
                    const tdCnaGpon = document.createElement('td')
                    tdCnaGpon.innerHTML = 'COLATINA'
                    const colCnaGpon = document.getElementById('cna')
                    colCnaGpon.append(tdCnaGpon)


               
                        // iniado
                    const tdIniCnaSlot10 = document.createElement('td')
                    tdIniCnaSlot10.innerHTML = gponCna.filter(slot10).filter(statusIniciada).length
                    const conIniCnaSlot10 = document.getElementById('cna')
                    conIniCnaSlot10.append(tdIniCnaSlot10)

                    const tdIniCnaSlot12 = document.createElement('td')
                    tdIniCnaSlot12.innerHTML = gponCna.filter(slot12).filter(statusIniciada).length
                    const conIniCnaSlot12 = document.getElementById('cna')
                    conIniCnaSlot12.append(tdIniCnaSlot12)

                    const tdIniCnaSlot15 = document.createElement('td')
                    tdIniCnaSlot15.innerHTML = gponCna.filter(slot15).filter(statusIniciada).length
                    const conIniCnaSlot15 = document.getElementById('cna')
                    conIniCnaSlot15.append(tdIniCnaSlot15)

                    const tdIniCnaSlot18 = document.createElement('td')
                    tdIniCnaSlot18.innerHTML = gponCna.filter(slot18).filter(statusIniciada).length
                    const conIniCnaSlot18 = document.getElementById('cna')
                    conIniCnaSlot18.append(tdIniCnaSlot18)
               

                    // não iniciada
                    
                    const tdNinCnaSlot10 = document.createElement('td')
                    tdNinCnaSlot10.innerHTML = gponCna.filter(slot10).filter(statusNaoIniciada).length
                    const conNinCnaSlot10 = document.getElementById('cna')
                    conNinCnaSlot10.append(tdNinCnaSlot10)

                    const tdNinCnaSlot12 = document.createElement('td')
                    tdNinCnaSlot12.innerHTML = gponCna.filter(slot12).filter(statusNaoIniciada).length
                    const conNinCnaSlot12 = document.getElementById('cna')
                    conNinCnaSlot12.append(tdNinCnaSlot12)

                    const tdNinCnaSlot15 = document.createElement('td')
                    tdNinCnaSlot15.innerHTML = gponCna.filter(slot15).filter(statusNaoIniciada).length
                    const conNinCnaSlot15 = document.getElementById('cna')
                    conNinCnaSlot15.append(tdNinCnaSlot15)

                    const tdNinCnaSlot18 = document.createElement('td')
                    tdNinCnaSlot18.innerHTML = gponCna.filter(slot18).filter(statusNaoIniciada).length
                    const conNinCnaSlot18 = document.getElementById('cna')
                    conNinCnaSlot18.append(tdNinCnaSlot18)



                    // GUARAPARI
                    const tdGriGpon = document.createElement('td')
                    tdGriGpon.innerHTML = 'GUARAPARI'
                    const colGriGpon = document.getElementById('gri')
                    colGriGpon.append(tdGriGpon)


                       // iniciado
                    const tdIniGriSlot10 = document.createElement('td')
                    tdIniGriSlot10.innerHTML = gponGri.filter(slot10).filter(statusIniciada).length
                    const conIniGriSlot10 = document.getElementById('gri')
                    conIniGriSlot10.append(tdIniGriSlot10)

                    const tdIniGriSlot12 = document.createElement('td')
                    tdIniGriSlot12.innerHTML = gponGri.filter(slot12).filter(statusIniciada).length
                    const conIniGriSlot12 = document.getElementById('gri')
                    conIniGriSlot12.append(tdIniGriSlot12)

                    const tdIniGriSlot15 = document.createElement('td')
                    tdIniGriSlot15.innerHTML = gponGri.filter(slot15).filter(statusIniciada).length
                    const conIniGriSlot15 = document.getElementById('gri')
                    conIniGriSlot15.append(tdIniGriSlot15)

                    const tdIniGriSlot18 = document.createElement('td')
                    tdIniGriSlot18.innerHTML = gponGri.filter(slot18).filter(statusIniciada).length
                    const conIniGriSlot18 = document.getElementById('gri')
                    conIniGriSlot18.append(tdIniGriSlot18)
               

                    // não iniciada
                    
                    const tdNinGriSlot10 = document.createElement('td')
                    tdNinGriSlot10.innerHTML = gponGri.filter(slot10).filter(statusNaoIniciada).length
                    const conNinGriSlot10 = document.getElementById('gri')
                    conNinGriSlot10.append(tdNinGriSlot10)

                    const tdNinGriSlot12 = document.createElement('td')
                    tdNinGriSlot12.innerHTML = gponGri.filter(slot12).filter(statusNaoIniciada).length
                    const conNinGriSlot12 = document.getElementById('gri')
                    conNinGriSlot12.append(tdNinGriSlot12)

                    const tdNinGriSlot15 = document.createElement('td')
                    tdNinGriSlot15.innerHTML = gponGri.filter(slot15).filter(statusNaoIniciada).length
                    const conNinGriSlot15 = document.getElementById('gri')
                    conNinGriSlot15.append(tdNinGriSlot15)

                    const tdNinGriSlot18 = document.createElement('td')
                    tdNinGriSlot18.innerHTML = gponGri.filter(slot18).filter(statusNaoIniciada).length
                    const conNinGriSlot18 = document.getElementById('gri')
                    conNinGriSlot18.append(tdNinGriSlot18)
               


                    // LINHARES
                    const tdLnsGpon = document.createElement('td')
                    tdLnsGpon.innerHTML = 'LINHARES'
                    const colLnsGpon = document.getElementById('lns')
                    colLnsGpon.append(tdLnsGpon)


               
                       // iniciado
                    const tdIniLnsSlot10 = document.createElement('td')
                    tdIniLnsSlot10.innerHTML = gponLns.filter(slot10).filter(statusIniciada).length
                    const conIniLnsSlot10 = document.getElementById('lns')
                    conIniLnsSlot10.append(tdIniLnsSlot10)

                    const tdIniLnsSlot12 = document.createElement('td')
                    tdIniLnsSlot12.innerHTML = gponLns.filter(slot12).filter(statusIniciada).length
                    const conIniLnsSlot12 = document.getElementById('lns')
                    conIniLnsSlot12.append(tdIniLnsSlot12)

                    const tdIniLnsSlot15 = document.createElement('td')
                    tdIniLnsSlot15.innerHTML = gponLns.filter(slot15).filter(statusIniciada).length
                    const conIniLnsSlot15 = document.getElementById('lns')
                    conIniLnsSlot15.append(tdIniLnsSlot15)

                    const tdIniLnsSlot18 = document.createElement('td')
                    tdIniLnsSlot18.innerHTML = gponLns.filter(slot18).filter(statusIniciada).length
                    const conIniLnsSlot18 = document.getElementById('lns')
                    conIniLnsSlot18.append(tdIniLnsSlot18)
               

                    // não iniciada
                    
                    const tdNinLnsSlot10 = document.createElement('td')
                    tdNinLnsSlot10.innerHTML = gponLns.filter(slot10).filter(statusNaoIniciada).length
                    const conNinLnsSlot10 = document.getElementById('lns')
                    conNinLnsSlot10.append(tdNinLnsSlot10)

                    const tdNinLnsSlot12 = document.createElement('td')
                    tdNinLnsSlot12.innerHTML = gponLns.filter(slot12).filter(statusNaoIniciada).length
                    const conNinLnsSlot12 = document.getElementById('lns')
                    conNinLnsSlot12.append(tdNinLnsSlot12)

                    const tdNinLnsSlot15 = document.createElement('td')
                    tdNinLnsSlot15.innerHTML = gponLns.filter(slot15).filter(statusNaoIniciada).length
                    const conNinLnsSlot15 = document.getElementById('lns')
                    conNinLnsSlot15.append(tdNinLnsSlot15)

                    const tdNinLnsSlot18 = document.createElement('td')
                    tdNinLnsSlot18.innerHTML = gponLns.filter(slot18).filter(statusNaoIniciada).length
                    const conNinLnsSlot18 = document.getElementById('lns')
                    conNinLnsSlot18.append(tdNinLnsSlot18)


                    // SANTA MARIA DE JETIBÁ

                    const tdSmjGpon = document.createElement('td')
                    tdSmjGpon.innerHTML = 'SANTA MARIA'
                    const colSmjGpon = document.getElementById('smj')
                    colSmjGpon.append(tdSmjGpon)


                    // iniciado
                    const tdIniSmjSlot10 = document.createElement('td')
                    tdIniSmjSlot10.innerHTML = gponSmj.filter(slot10).filter(statusIniciada).length
                    const conIniSmjSlot10 = document.getElementById('smj')
                    conIniSmjSlot10.append(tdIniSmjSlot10)

                    const tdIniSmjSlot12 = document.createElement('td')
                    tdIniSmjSlot12.innerHTML = gponSmj.filter(slot12).filter(statusIniciada).length
                    const conIniSmjSlot12 = document.getElementById('smj')
                    conIniSmjSlot12.append(tdIniSmjSlot12)

                    const tdIniSmjSlot15 = document.createElement('td')
                    tdIniSmjSlot15.innerHTML = gponSmj.filter(slot15).filter(statusIniciada).length
                    const conIniSmjSlot15 = document.getElementById('smj')
                    conIniSmjSlot15.append(tdIniSmjSlot15)

                    const tdIniSmjSlot18 = document.createElement('td')
                    tdIniSmjSlot18.innerHTML = gponSmj.filter(slot18).filter(statusIniciada).length
                    const conIniSmjSlot18 = document.getElementById('smj')
                    conIniSmjSlot18.append(tdIniSmjSlot18)
               

                    // não iniciada
                    
                    const tdNinSmjSlot10 = document.createElement('td')
                    tdNinSmjSlot10.innerHTML = gponSmj.filter(slot10).filter(statusNaoIniciada).length
                    const conNinSmjSlot10 = document.getElementById('smj')
                    conNinSmjSlot10.append(tdNinSmjSlot10)

                    const tdNinSmjSlot12 = document.createElement('td')
                    tdNinSmjSlot12.innerHTML = gponSmj.filter(slot12).filter(statusNaoIniciada).length
                    const conNinSmjSlot12 = document.getElementById('smj')
                    conNinSmjSlot12.append(tdNinSmjSlot12)

                    const tdNinSmjSlot15 = document.createElement('td')
                    tdNinSmjSlot15.innerHTML = gponSmj.filter(slot15).filter(statusNaoIniciada).length
                    const conNinSmjSlot15 = document.getElementById('smj')
                    conNinSmjSlot15.append(tdNinSmjSlot15)

                    const tdNinSmjSlot18 = document.createElement('td')
                    tdNinSmjSlot18.innerHTML = gponSmj.filter(slot18).filter(statusNaoIniciada).length
                    const conNinSmjSlot18 = document.getElementById('smj')
                    conNinSmjSlot18.append(tdNinSmjSlot18)



                    

                    // SÃO MATEUS
                    const tdSmtGpon = document.createElement('td')
                    tdSmtGpon.innerHTML = 'SÃO MATEUS'
                    const colSmtGpon = document.getElementById('smt')
                    colSmtGpon.append(tdSmtGpon)


                    
                    // iniciado
                    const tdIniSmtSlot10 = document.createElement('td')
                    tdIniSmtSlot10.innerHTML = gponSmt.filter(slot10).filter(statusIniciada).length
                    const conIniSmtSlot10 = document.getElementById('smt')
                    conIniSmtSlot10.append(tdIniSmtSlot10)

                    const tdIniSmtSlot12 = document.createElement('td')
                    tdIniSmtSlot12.innerHTML = gponSmt.filter(slot12).filter(statusIniciada).length
                    const conIniSmtSlot12 = document.getElementById('smt')
                    conIniSmtSlot12.append(tdIniSmtSlot12)

                    const tdIniSmtSlot15 = document.createElement('td')
                    tdIniSmtSlot15.innerHTML = gponSmt.filter(slot15).filter(statusIniciada).length
                    const conIniSmtSlot15 = document.getElementById('smt')
                    conIniSmtSlot15.append(tdIniSmtSlot15)

                    const tdIniSmtSlot18 = document.createElement('td')
                    tdIniSmtSlot18.innerHTML = gponSmt.filter(slot18).filter(statusIniciada).length
                    const conIniSmtSlot18 = document.getElementById('smt')
                    conIniSmtSlot18.append(tdIniSmtSlot18)
               

                    // não iniciada
                    
                    const tdNinSmtSlot10 = document.createElement('td')
                    tdNinSmtSlot10.innerHTML = gponSmt.filter(slot10).filter(statusNaoIniciada).length
                    const conNinSmtSlot10 = document.getElementById('smt')
                    conNinSmtSlot10.append(tdNinSmtSlot10)

                    const tdNinSmtSlot12 = document.createElement('td')
                    tdNinSmtSlot12.innerHTML = gponSmt.filter(slot12).filter(statusNaoIniciada).length
                    const conNinSmtSlot12 = document.getElementById('smt')
                    conNinSmtSlot12.append(tdNinSmtSlot12)

                    const tdNinSmtSlot15 = document.createElement('td')
                    tdNinSmtSlot15.innerHTML = gponSmt.filter(slot15).filter(statusNaoIniciada).length
                    const conNinSmtSlot15 = document.getElementById('smt')
                    conNinSmtSlot15.append(tdNinSmtSlot15)

                    const tdNinSmtSlot18 = document.createElement('td')
                    tdNinSmtSlot18.innerHTML = gponSmt.filter(slot18).filter(statusNaoIniciada).length
                    const conNinSmtSlot18 = document.getElementById('smt')
                    conNinSmtSlot18.append(tdNinSmtSlot18)



                    // SERRA
                    const tdSeaGpon = document.createElement('td')
                    tdSeaGpon.innerHTML = 'SERRA'
                    const colSeaGpon = document.getElementById('sea')
                    colSeaGpon.append(tdSeaGpon)


                    // iniciado
                    const tdIniSeaSlot10 = document.createElement('td')
                    tdIniSeaSlot10.innerHTML = gponSea.filter(slot10).filter(statusIniciada).length
                    const conIniSeaSlot10 = document.getElementById('sea')
                    conIniSeaSlot10.append(tdIniSeaSlot10)

                    const tdIniSeaSlot12 = document.createElement('td')
                    tdIniSeaSlot12.innerHTML = gponSea.filter(slot12).filter(statusIniciada).length
                    const conIniSeaSlot12 = document.getElementById('sea')
                    conIniSeaSlot12.append(tdIniSeaSlot12)

                    const tdIniSeaSlot15 = document.createElement('td')
                    tdIniSeaSlot15.innerHTML = gponSea.filter(slot15).filter(statusIniciada).length
                    const conIniSeaSlot15 = document.getElementById('sea')
                    conIniSeaSlot15.append(tdIniSeaSlot15)

                    const tdIniSeaSlot18 = document.createElement('td')
                    tdIniSeaSlot18.innerHTML = gponSea.filter(slot18).filter(statusIniciada).length
                    const conIniSeaSlot18 = document.getElementById('sea')
                    conIniSeaSlot18.append(tdIniSeaSlot18)
               

                    // não iniciada
                    
                    const tdNinSeaSlot10 = document.createElement('td')
                    tdNinSeaSlot10.innerHTML = gponSea.filter(slot10).filter(statusNaoIniciada).length
                    const conNinSeaSlot10 = document.getElementById('sea')
                    conNinSeaSlot10.append(tdNinSeaSlot10)

                    const tdNinSeaSlot12 = document.createElement('td')
                    tdNinSeaSlot12.innerHTML = gponSea.filter(slot12).filter(statusNaoIniciada).length
                    const conNinSeaSlot12 = document.getElementById('sea')
                    conNinSeaSlot12.append(tdNinSeaSlot12)

                    const tdNinSeaSlot15 = document.createElement('td')
                    tdNinSeaSlot15.innerHTML = gponSea.filter(slot15).filter(statusNaoIniciada).length
                    const conNinSeaSlot15 = document.getElementById('sea')
                    conNinSeaSlot15.append(tdNinSeaSlot15)

                    const tdNinSeaSlot18 = document.createElement('td')
                    tdNinSeaSlot18.innerHTML = gponSea.filter(slot18).filter(statusNaoIniciada).length
                    const conNinSeaSlot18 = document.getElementById('sea')
                    conNinSeaSlot18.append(tdNinSeaSlot18)


                    

                    // VILA VELHA
                    const tdVvaGpon = document.createElement('td')
                    tdVvaGpon.innerHTML = 'VILA VELHA'
                    const colVvaGpon = document.getElementById('vva')
                    colVvaGpon.append(tdVvaGpon)



                      // iniciado
                    const tdIniVvaSlot10 = document.createElement('td')
                    tdIniVvaSlot10.innerHTML = gponVva.filter(slot10).filter(statusIniciada).length
                    const conIniVvaSlot10 = document.getElementById('vva')
                    conIniVvaSlot10.append(tdIniVvaSlot10)

                    const tdIniVvaSlot12 = document.createElement('td')
                    tdIniVvaSlot12.innerHTML = gponVva.filter(slot12).filter(statusIniciada).length
                    const conIniVvaSlot12 = document.getElementById('vva')
                    conIniVvaSlot12.append(tdIniVvaSlot12)

                    const tdIniVvaSlot15 = document.createElement('td')
                    tdIniVvaSlot15.innerHTML = gponVva.filter(slot15).filter(statusIniciada).length
                    const conIniVvaSlot15 = document.getElementById('vva')
                    conIniVvaSlot15.append(tdIniVvaSlot15)

                    const tdIniVvaSlot18 = document.createElement('td')
                    tdIniVvaSlot18.innerHTML = gponVva.filter(slot18).filter(statusIniciada).length
                    const conIniVvaSlot18 = document.getElementById('vva')
                    conIniVvaSlot18.append(tdIniVvaSlot18)
               

                    // não iniciada
                    
                    const tdNinVvaSlot10 = document.createElement('td')
                    tdNinVvaSlot10.innerHTML = gponVva.filter(slot10).filter(statusNaoIniciada).length
                    const conNinVvaSlot10 = document.getElementById('vva')
                    conNinVvaSlot10.append(tdNinVvaSlot10)

                    const tdNinVvaSlot12 = document.createElement('td')
                    tdNinVvaSlot12.innerHTML = gponVva.filter(slot12).filter(statusNaoIniciada).length
                    const conNinVvaSlot12 = document.getElementById('vva')
                    conNinVvaSlot12.append(tdNinVvaSlot12)

                    const tdNinVvaSlot15 = document.createElement('td')
                    tdNinVvaSlot15.innerHTML = gponVva.filter(slot15).filter(statusNaoIniciada).length
                    const conNinVvaSlot15 = document.getElementById('vva')
                    conNinVvaSlot15.append(tdNinVvaSlot15)

                    const tdNinVvaSlot18 = document.createElement('td')
                    tdNinVvaSlot18.innerHTML = gponVva.filter(slot18).filter(statusNaoIniciada).length
                    const conNinVvaSlot18 = document.getElementById('vva')
                    conNinVvaSlot18.append(tdNinVvaSlot18)




                    // VITORIA
                    const tdVtaGpon = document.createElement('td')
                    tdVtaGpon.innerHTML = 'VITÓRIA'
                    const colVtaGpon = document.getElementById('vta')
                    colVtaGpon.append(tdVtaGpon)


                    
                      // iniciado
                    const tdIniVtaSlot10 = document.createElement('td')
                    tdIniVtaSlot10.innerHTML = gponVta.filter(slot10).filter(statusIniciada).length
                    const conIniVtaSlot10 = document.getElementById('vta')
                    conIniVtaSlot10.append(tdIniVtaSlot10)

                    const tdIniVtaSlot12 = document.createElement('td')
                    tdIniVtaSlot12.innerHTML = gponVta.filter(slot12).filter(statusIniciada).length
                    const conIniVtaSlot12 = document.getElementById('vta')
                    conIniVtaSlot12.append(tdIniVtaSlot12)

                    const tdIniVtaSlot15 = document.createElement('td')
                    tdIniVtaSlot15.innerHTML = gponVta.filter(slot15).filter(statusIniciada).length
                    const conIniVtaSlot15 = document.getElementById('vta')
                    conIniVtaSlot15.append(tdIniVtaSlot15)

                    const tdIniVtaSlot18 = document.createElement('td')
                    tdIniVtaSlot18.innerHTML = gponVta.filter(slot18).filter(statusIniciada).length
                    const conIniVtaSlot18 = document.getElementById('vta')
                    conIniVtaSlot18.append(tdIniVtaSlot18)
               

                    // não iniciada
                    
                    const tdNinVtaSlot10 = document.createElement('td')
                    tdNinVtaSlot10.innerHTML = gponVta.filter(slot10).filter(statusNaoIniciada).length
                    const conNinVtaSlot10 = document.getElementById('vta')
                    conNinVtaSlot10.append(tdNinVtaSlot10)

                    const tdNinVtaSlot12 = document.createElement('td')
                    tdNinVtaSlot12.innerHTML = gponVta.filter(slot12).filter(statusNaoIniciada).length
                    const conNinVtaSlot12 = document.getElementById('vta')
                    conNinVtaSlot12.append(tdNinVtaSlot12)

                    const tdNinVtaSlot15 = document.createElement('td')
                    tdNinVtaSlot15.innerHTML = gponVta.filter(slot15).filter(statusNaoIniciada).length
                    const conNinVtaSlot15 = document.getElementById('vta')
                    conNinVtaSlot15.append(tdNinVtaSlot15)

                    const tdNinVtaSlot18 = document.createElement('td')
                    tdNinVtaSlot18.innerHTML = gponVta.filter(slot18).filter(statusNaoIniciada).length
                    const conNinVtaSlot18 = document.getElementById('vta')
                    conNinVtaSlot18.append(tdNinVtaSlot18)

                    

                    // VIANA
                    const tdViaGpon = document.createElement('td')
                    tdViaGpon.innerHTML = 'VIANA'
                    const colViaGpon = document.getElementById('via')
                    colViaGpon.append(tdViaGpon)

               


                     // iniciado
                    const tdIniViaSlot10 = document.createElement('td')
                    tdIniViaSlot10.innerHTML = gponVia.filter(slot10).filter(statusIniciada).length
                    const conIniViaSlot10 = document.getElementById('via')
                    conIniViaSlot10.append(tdIniViaSlot10)

                    const tdIniViaSlot12 = document.createElement('td')
                    tdIniViaSlot12.innerHTML = gponVia.filter(slot12).filter(statusIniciada).length
                    const conIniViaSlot12 = document.getElementById('via')
                    conIniViaSlot12.append(tdIniViaSlot12)

                    const tdIniViaSlot15 = document.createElement('td')
                    tdIniViaSlot15.innerHTML = gponVia.filter(slot15).filter(statusIniciada).length
                    const conIniViaSlot15 = document.getElementById('via')
                    conIniViaSlot15.append(tdIniViaSlot15)

                    const tdIniViaSlot18 = document.createElement('td')
                    tdIniViaSlot18.innerHTML = gponVia.filter(slot18).filter(statusIniciada).length
                    const conIniViaSlot18 = document.getElementById('via')
                    conIniViaSlot18.append(tdIniViaSlot18)
               

                    // não iniciada
                    
                    const tdNinViaSlot10 = document.createElement('td')
                    tdNinViaSlot10.innerHTML = gponVia.filter(slot10).filter(statusNaoIniciada).length
                    const conNinViaSlot10 = document.getElementById('via')
                    conNinViaSlot10.append(tdNinViaSlot10)

                    const tdNinViaSlot12 = document.createElement('td')
                    tdNinViaSlot12.innerHTML = gponVia.filter(slot12).filter(statusNaoIniciada).length
                    const conNinViaSlot12 = document.getElementById('via')
                    conNinViaSlot12.append(tdNinViaSlot12)

                    const tdNinViaSlot15 = document.createElement('td')
                    tdNinViaSlot15.innerHTML = gponVia.filter(slot15).filter(statusNaoIniciada).length
                    const conNinViaSlot15 = document.getElementById('via')
                    conNinViaSlot15.append(tdNinViaSlot15)

                    const tdNinViaSlot18 = document.createElement('td')
                    tdNinViaSlot18.innerHTML = gponVia.filter(slot18).filter(statusNaoIniciada).length
                    const conNinViaSlot18 = document.getElementById('via')
                    conNinViaSlot18.append(tdNinViaSlot18)
                                   


                    // TOTAL
                              

               
                    const tdGpon = document.createElement('td')
                    tdGpon.className = 'tdGpon'
                    tdGpon.innerHTML = 'TOTAL'
                    const colGpon = document.getElementById('total')
                    colGpon.append(tdGpon)      

                    // iniado

                    const tdIniSlot10 = document.createElement('td')
                    tdIniSlot10.innerHTML = dataGpon.filter(slot10).filter(statusIniciada).length
                    const conIniSlot10 = document.getElementById('total')
                    conIniSlot10.append(tdIniSlot10)

                    const tdIniSlot12 = document.createElement('td')
                    tdIniSlot12.innerHTML = dataGpon.filter(slot12).filter(statusIniciada).length
                    const conIniSlot12 = document.getElementById('total')
                    conIniSlot12.append(tdIniSlot12)

                    const tdIniSlot15 = document.createElement('td')
                    tdIniSlot15.innerHTML = dataGpon.filter(slot15).filter(statusIniciada).length
                    const conIniSlot15 = document.getElementById('total')
                    conIniSlot15.append(tdIniSlot15)

                    const tdIniSlot18 = document.createElement('td')
                    tdIniSlot18.innerHTML = dataGpon.filter(slot18).filter(statusIniciada).length
                    const conIniSlot18 = document.getElementById('total')
                    conIniSlot18.append(tdIniSlot18)
               

                    // não iniciada
                    
                    const tdNinSlot10 = document.createElement('td')
                    tdNinSlot10.innerHTML = dataGpon.filter(slot10).filter(statusNaoIniciada).length
                    const conNinSlot10 = document.getElementById('total')
                    conNinSlot10.append(tdNinSlot10)

                    const tdNinSlot12 = document.createElement('td')
                    tdNinSlot12.innerHTML = dataGpon.filter(slot12).filter(statusNaoIniciada).length
                    const conNinSlot12 = document.getElementById('total')
                    conNinSlot12.append(tdNinSlot12)

                    const tdNinSlot15 = document.createElement('td')
                    tdNinSlot15.innerHTML = dataGpon.filter(slot15).filter(statusNaoIniciada).length
                    const conNinSlot15 = document.getElementById('total')
                    conNinSlot15.append(tdNinSlot15)

                    const tdNinSlot18 = document.createElement('td')
                    tdNinSlot18.innerHTML = dataGpon.filter(slot18).filter(statusNaoIniciada).length
                    const conNinSlot18 = document.getElementById('total')
                    conNinSlot18.append(tdNinSlot18)
                    


                    // ============= btn Download

                    const btnProducao = document.createElement('button')
                    btnProducao.id = 'btnDonwload'
                    btnProducao.innerHTML = 'BAIXAR IMAGEM'
                    const btnProducao2 = document.getElementById('div-download')
                    btnProducao2.append(btnProducao)
               
          
                    let btnGenerator = document.querySelector('#btnDonwload')
                    let btnDownload = document.querySelector('.download')

                    btnGenerator.addEventListener('click', () => {
                         html2canvas(document.querySelector("#canvasDown")).then(canvas => {
                              document.body.appendChild(canvas)
                              btnDownload.href = canvas.toDataURL('image/png');
                              btnDownload.download = 'producao';
                              btnDownload.click();
                         });
                    })



               });



          }
     }


})
