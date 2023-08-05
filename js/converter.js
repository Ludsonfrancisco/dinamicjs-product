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
                                                       item["Tipo de Atividade"] === 'Defeito Banda Larga' ||
                                                       item["Tipo de Atividade"] === 'Defeito Banda Larga' 
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


               // ====CREATE CABEÇALHO TABELA PRODUÇAO GPON ====
                    // const titleProducao = document.createElement('h1')
                    // titleProducao.className = 'title-producao'
                    // titleProducao.innerHTML = 'PRODUÇÃO'
                    // const tProducao = document.getElementById("producao")
                    // tProducao.appendChild(titleProducao)



// ''''                         <span class="span-title" style="
//                               background: #767CFF;
//                               width: auto;
//                               height: 2rem;
//                               /* color: #767CFF; */
//                               font-size: 20px;
//                               font-weight: bold;
//                               border-radius: 3px;
//                               text-align: center;
//                               ">GPON</span>
// ''''

                    const titleGpon = document.createElement('span')
                    titleGpon.innerHTML = 'GPON'
                    const tGpon = document.getElementById("title-gpon")
                    tGpon.append(titleGpon)

                    const tdCidade = document.createElement('td')
                    tdCidade.className = 'tdCidade'
                    tdCidade.innerHTML = 'CIDADE'

                    const tdConcluida = document.createElement('td')
                    tdConcluida.className = 'tdConcluida' 
                    tdConcluida.innerHTML = 'CONCLUIDA'

                    const tdIniciada = document.createElement('td')
                    tdIniciada.className = 'tdIniciada'
                    tdIniciada.innerHTML = 'INICIADA'

                    const tdNaoiniciada = document.createElement('td')
                    tdNaoiniciada.className = 'tdNin'
                    tdNaoiniciada.innerHTML = 'NÃO INICIADA'

                    const total = document.createElement('td')
                    total.className = 'tdTotal'
                    total.innerHTML = 'TOTAL'

                    const tabela = document.getElementById('cabecalho')
                    tabela.append(tdCidade)
                 ""   tabela.append(tdIniciada)
                    tabela.append(tdNaoiniciada)
                    tabela.append(total)

                    // ======== CREATE CIDADE GPON ====


                    // ARACRUZ
                    const tdArcGpon = document.createElement('td')
                    tdArcGpon.innerHTML = 'ARACRUZ'
                    const colArcGpon = document.getElementById('arc')
                    colArcGpon.append(tdArcGpon)


                    const tdConArcGpon = document.createElement('td')
                    tdConArcGpon.innerHTML = gponArc.filter(statusConcluida).length
                    const colConArcGpon = document.getElementById('arc')
                    colConArcGpon.append(tdConArcGpon)

                    const tdIniArcGpon = document.createElement('td')
                    tdIniArcGpon.innerHTML = gponArc.filter(statusIniciada).length
                    const conIniArcGpon = document.getElementById('arc')
                    conIniArcGpon.append(tdIniArcGpon)

                    const tdNinArcGpon = document.createElement('td')
                    tdNinArcGpon.innerHTML = gponArc.filter(statusNaoIniciada).length
                    const colNinArcGpon = document.getElementById('arc')
                    colNinArcGpon.append(tdNinArcGpon)

                    const tdTotalArcGpon = document.createElement('td')
                    tdTotalArcGpon.innerHTML = gponArc.filter(conIniNin).length
                    const colTotalArcGpon = document.getElementById('arc')
                    colTotalArcGpon.append(tdTotalArcGpon)


                    // CACHOEIRO
                    const tdCimGpon = document.createElement('td')
                    tdCimGpon.innerHTML = 'CACHOEIRO'
                    const colCimGpon = document.getElementById('cim')
                    colCimGpon.append(tdCimGpon)


                    const tdConCimGpon = document.createElement('td')
                    tdConCimGpon.innerHTML = gponCim.filter(statusConcluida).length
                    const colConCimGpon = document.getElementById('cim')
                    colConCimGpon.append(tdConCimGpon)

                    const tdIniCimGpon = document.createElement('td')
                    tdIniCimGpon.innerHTML = gponCim.filter(statusIniciada).length
                    const colIniCimGpon = document.getElementById('cim')
                    colIniCimGpon.append(tdIniCimGpon)

                    const tdNinCimGPon = document.createElement('td')
                    tdNinCimGPon.innerHTML = gponCim.filter(statusNaoIniciada).length
                    const colNinCimGpon = document.getElementById('cim')
                    colNinCimGpon.append(tdNinCimGPon)

                    const tdTotalCimGpon = document.createElement('td')
                    tdTotalCimGpon.innerHTML = gponCim.filter(conIniNin).length
                    const colTotalCimGpon = document.getElementById('cim')
                    colTotalCimGpon.append(tdTotalCimGpon)


                    // CARIACICA
                    const tdCcaGpon = document.createElement('td')
                    tdCcaGpon.innerHTML = 'CARIACICA'
                    const colCcaGpon = document.getElementById('cca')
                    colCcaGpon.append(tdCcaGpon)


                    const tdConCcaGpon = document.createElement('td')
                    tdConCcaGpon.innerHTML = gponCca.filter(statusConcluida).length
                    const colConCcaGpon = document.getElementById('cca')
                    colConCcaGpon.append(tdConCcaGpon)

                    const tdIniCcaGpon = document.createElement('td')
                    tdIniCcaGpon.innerHTML = gponCca.filter(statusIniciada).length
                    const colIniCcaGpon = document.getElementById('cca')
                    colIniCcaGpon.append(tdIniCcaGpon)

                    const tdNinCcaGpon = document.createElement('td')
                    tdNinCcaGpon.innerHTML = gponCca.filter(statusNaoIniciada).length
                    const colNinCcaGpon = document.getElementById('cca')
                    colNinCcaGpon.append(tdNinCcaGpon)

                    const tdTotalCcaGpon = document.createElement('td')
                    tdTotalCcaGpon.innerHTML = gponCca.filter(conIniNin).length
                    const colTotalCcaGpon = document.getElementById('cca')
                    colTotalCcaGpon.append(tdTotalCcaGpon)


                    // COLATINA
                    const tdCnaGpon = document.createElement('td')
                    tdCnaGpon.innerHTML = 'COLATINA'
                    const colCnaGpon = document.getElementById('cna')
                    colCnaGpon.append(tdCnaGpon)

                    const tdConCnaGpon = document.createElement('td')
                    tdConCnaGpon.innerHTML = gponCna.filter(statusConcluida).length
                    const colConCnaGpon = document.getElementById('cna')
                    colConCnaGpon.append(tdConCnaGpon)

                    const tdIniCnaGpon = document.createElement('td')
                    tdIniCnaGpon.innerHTML = gponCna.filter(statusIniciada).length
                    const colIniCnaGpon = document.getElementById('cna')
                    colIniCnaGpon.append(tdIniCnaGpon)

                    const tdNinCnaGpon = document.createElement('td')
                    tdNinCnaGpon.innerHTML = gponCna.filter(statusNaoIniciada).length
                    const colNinCnaGpon = document.getElementById('cna')
                    colNinCnaGpon.append(tdNinCnaGpon)

                    const tdTotalCnaGpon = document.createElement('td')
                    tdTotalCnaGpon.innerHTML = gponCna.filter(conIniNin).length
                    const colTotalCnaGpon = document.getElementById('cna')
                    colTotalCnaGpon.append(tdTotalCnaGpon)


                    // GUARAPARI
                    const tdGriGpon = document.createElement('td')
                    tdGriGpon.innerHTML = 'GUARAPARI'
                    const colGriGpon = document.getElementById('gri')
                    colGriGpon.append(tdGriGpon)

                    const tdConGriGpon = document.createElement('td')
                    tdConGriGpon.innerHTML = gponGri.filter(statusConcluida).length
                    const colConGriGpon = document.getElementById('gri')
                    colConGriGpon.append(tdConGriGpon)

                    const tdIniGriGpon = document.createElement('td')
                    tdIniGriGpon.innerHTML = gponGri.filter(statusIniciada).length
                    const colIniGriGpon = document.getElementById('gri')
                    colIniGriGpon.append(tdIniGriGpon)

                    const tdNinGriGpon = document.createElement('td')
                    tdNinGriGpon.innerHTML = gponGri.filter(statusNaoIniciada).length
                    const colNinGriGpon = document.getElementById('gri')
                    colNinGriGpon.append(tdNinGriGpon)

                    const tdTotalGriGpon = document.createElement('td')
                    tdTotalGriGpon.innerHTML = gponGri.filter(conIniNin).length
                    const colTotalGriGpon = document.getElementById('gri')
                    colTotalGriGpon.append(tdTotalGriGpon)


                    // LINHARES
                    const tdLnsGpon = document.createElement('td')
                    tdLnsGpon.innerHTML = 'LINHARES'
                    const colLnsGpon = document.getElementById('lns')
                    colLnsGpon.append(tdLnsGpon)

                    const tdConLnsGpon = document.createElement('td')
                    tdConLnsGpon.innerHTML = gponLns.filter(statusConcluida).length
                    const colConLnsGpon = document.getElementById('lns')
                    colConLnsGpon.append(tdConLnsGpon)

                    const tdIniLnsGpon = document.createElement('td')
                    tdIniLnsGpon.innerHTML = gponLns.filter(statusIniciada).length
                    const colIniLnsGpon = document.getElementById('lns')
                    colIniLnsGpon.append(tdIniLnsGpon)

                    const tdNinLnsGpon = document.createElement('td')
                    tdNinLnsGpon.innerHTML = gponLns.filter(statusNaoIniciada).length
                    const colNinLnsGpon = document.getElementById('lns')
                    colNinLnsGpon.append(tdNinLnsGpon)

                    const tdTotalLnsGpon = document.createElement('td')
                    tdTotalLnsGpon.innerHTML = gponLns.filter(conIniNin).length
                    const colTotalLnsGpon = document.getElementById('lns')
                    colTotalLnsGpon.append(tdTotalLnsGpon)



                    // SANTA MARIA DE JETIBÁ

                    const tdSmjGpon = document.createElement('td')
                    tdSmjGpon.innerHTML = 'SANTA MARIA'
                    const colSmjGpon = document.getElementById('smj')
                    colSmjGpon.append(tdSmjGpon)

                    const tdConSmjGpon = document.createElement('td')
                    tdConSmjGpon.innerHTML = gponSmj.filter(statusConcluida).length
                    const colConSmjGpon = document.getElementById('smj')
                    colConSmjGpon.append(tdConSmjGpon)

                    const tdIniSmjGpon = document.createElement('td')
                    tdIniSmjGpon.innerHTML = gponSmj.filter(statusIniciada).length
                    const colIniSmjGpon = document.getElementById('smj')
                    colIniSmjGpon.append(tdIniSmjGpon)

                    const tdNinSmjGpon = document.createElement('td')
                    tdNinSmjGpon.innerHTML = gponSmj.filter(statusNaoIniciada).length
                    const colNinSmjGpon = document.getElementById('smj')
                    colNinSmjGpon.append(tdNinSmjGpon)

                    const tdTotalSmjGpon = document.createElement('td')
                    tdTotalSmjGpon.innerHTML = gponSmj.filter(conIniNin).length
                    const colTotalSmjGpon = document.getElementById('smj')
                    colTotalSmjGpon.append(tdTotalSmjGpon)



                    // SÃO MATEUS
                    const tdSmtGpon = document.createElement('td')
                    tdSmtGpon.innerHTML = 'SÃO MATEUS'
                    const colSmtGpon = document.getElementById('smt')
                    colSmtGpon.append(tdSmtGpon)

                    const tdConSmtGpon = document.createElement('td')
                    tdConSmtGpon.innerHTML = gponSmt.filter(statusConcluida).length
                    const colConSmtGpon = document.getElementById('smt')
                    colConSmtGpon.append(tdConSmtGpon)

                    const tdIniSmtGpon = document.createElement('td')
                    tdIniSmtGpon.innerHTML = gponSmt.filter(statusIniciada).length
                    const colIniSmtGpon = document.getElementById('smt')
                    colIniSmtGpon.append(tdIniSmtGpon)

                    const tdNinSmtGpon = document.createElement('td')
                    tdNinSmtGpon.innerHTML = gponSmt.filter(statusNaoIniciada).length
                    const colNinSmtGpon = document.getElementById('smt')
                    colNinSmtGpon.append(tdNinSmtGpon)

                    const tdTotalSmtGpon = document.createElement('td')
                    tdTotalSmtGpon.innerHTML = gponSmt.filter(conIniNin).length
                    const colTotalSmtGpon = document.getElementById('smt')
                    colTotalSmtGpon.append(tdTotalSmtGpon)

                    // SERRA
                    const tdSeaGpon = document.createElement('td')
                    tdSeaGpon.innerHTML = 'SERRA'
                    const colSeaGpon = document.getElementById('sea')
                    colSeaGpon.append(tdSeaGpon)


                    const tdConSeaGpon = document.createElement('td')
                    tdConSeaGpon.innerHTML = gponSea.filter(statusConcluida).length
                    const colConSeaGpon = document.getElementById('sea')
                    colConSeaGpon.append(tdConSeaGpon)

                    const tdIniSeaGpon = document.createElement('td')
                    tdIniSeaGpon.innerHTML = gponSea.filter(statusIniciada).length
                    const colIniSeaGpon = document.getElementById('sea')
                    colIniSeaGpon.append(tdIniSeaGpon)

                    const tdNinSeaGpon = document.createElement('td')
                    tdNinSeaGpon.innerHTML = gponSea.filter(statusNaoIniciada).length
                    const colNinSeaGpon = document.getElementById('sea')
                    colNinSeaGpon.append(tdNinSeaGpon)

                    const tdTotalSea = document.createElement('td')
                    tdTotalSea.innerHTML = gponSea.filter(conIniNin).length
                    const colTotalSea = document.getElementById('sea')
                    colTotalSea.append(tdTotalSea)


                    // VILA VELHA
                    const tdVvaGpon = document.createElement('td')
                    tdVvaGpon.innerHTML = 'VILA VELHA'
                    const colVvaGpon = document.getElementById('vva')
                    colVvaGpon.append(tdVvaGpon)

                    const tdConVvaGpon = document.createElement('td')
                    tdConVvaGpon.innerHTML = gponVva.filter(statusConcluida).length
                    const colConVva = document.getElementById('vva')
                    colConVva.append(tdConVvaGpon)

                    const tdIniVvaGpon = document.createElement('td')
                    tdIniVvaGpon.innerHTML = gponVva.filter(statusIniciada).length
                    const colIniVva = document.getElementById('vva')
                    colIniVva.append(tdIniVvaGpon)

                    const tdNinVvaGpon = document.createElement('td')
                    tdNinVvaGpon.innerHTML = gponVva.filter(statusNaoIniciada).length
                    const colNinVvaGpon = document.getElementById('vva')
                    colNinVvaGpon.append(tdNinVvaGpon)

                    const tdTotalVvaGpon = document.createElement('td')
                    tdTotalVvaGpon.innerHTML = gponVva.filter(conIniNin).length
                    const colTotalVvaGpon = document.getElementById('vva')
                    colTotalVvaGpon.append(tdTotalVvaGpon)


                    // VITORIA
                    const tdVtaGpon = document.createElement('td')
                    tdVtaGpon.innerHTML = 'VITÓRIA'
                    const colVtaGpon = document.getElementById('vta')
                    colVtaGpon.append(tdVtaGpon)

                    const tdConVtaGpon = document.createElement('td')
                    tdConVtaGpon.innerHTML = gponVta.filter(statusConcluida).length
                    const colConVtaGpon = document.getElementById('vta')
                    colConVtaGpon.append(tdConVtaGpon)

                    const tdIniVtaGpon = document.createElement('td')
                    tdIniVtaGpon.innerHTML = gponVta.filter(statusIniciada).length
                    const colIniVtaGpon = document.getElementById('vta')
                    colIniVtaGpon.append(tdIniVtaGpon)

                    const tdNinVtaGpon = document.createElement('td')
                    tdNinVtaGpon.innerHTML = gponVta.filter(statusNaoIniciada).length
                    const colNinVtaGpon = document.getElementById('vta')
                    colNinVtaGpon.append(tdNinVtaGpon)

                    const tdTotalVta = document.createElement('td')
                    tdTotalVta.innerHTML = gponVta.filter(conIniNin).length
                    const colTotalVta = document.getElementById('vta')
                    colTotalVta.append(tdTotalVta)


                    // VIANA
                    const tdViaGpon = document.createElement('td')
                    tdViaGpon.innerHTML = 'VIANA'
                    const colViaGpon = document.getElementById('via')
                    colViaGpon.append(tdViaGpon)

                    const tdConViaGpon = document.createElement('td')
                    tdConViaGpon.innerHTML = gponVia.filter(statusConcluida).length
                    const colConViaGpon = document.getElementById('via')
                    colConViaGpon.append(tdConViaGpon)

                    const tdIniViaGpon = document.createElement('td')
                    tdIniViaGpon.innerHTML = gponVia.filter(statusIniciada).length
                    const colIniViaGpon = document.getElementById('via')
                    colIniViaGpon.append(tdIniViaGpon)

                    const tdNinViaGpon = document.createElement('td')
                    tdNinViaGpon.innerHTML = gponVia.filter(statusNaoIniciada).length
                    const colNinViaGpon = document.getElementById('via')
                    colNinViaGpon.append(tdNinViaGpon)

                    const tdTotalVia = document.createElement('td')
                    tdTotalVia.innerHTML = gponVia.filter(conIniNin).length
                    const colTotalVia = document.getElementById('via')
                    colTotalVia.append(tdTotalVia)



                                   


                    // TOTAL
                    const tdGpon = document.createElement('td')
                    tdGpon.className = 'tdGpon'
                    tdGpon.innerHTML = 'TOTAL'
                    const colGpon = document.getElementById('total')
                    colGpon.append(tdGpon)

                    
                    const tdIniGpon = document.createElement('td')
                    tdIniGpon.innerHTML = dataGpon.filter(statusIniciada).length
                    const colIniGpon = document.getElementById('total')
                    colIniGpon.append(tdIniGpon)

                    const tdNinGpon = document.createElement('td')
                    tdNinGpon.innerHTML = dataGpon.filter(statusNaoIniciada).length 
                    const colNinGpon = document.getElementById('total')
                    colNinGpon.append(tdNinGpon)

                    const tdSumGpon = document.createElement('td')
                    tdSumGpon.innerHTML = dataGpon.filter(conIniNin).length
                    const colSumGpon = document.getElementById('total')
                    colSumGpon.append(tdSumGpon)


               



               
          
     
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





