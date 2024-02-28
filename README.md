# AGENDA-X-PLANILHA
Integração desenvolvida para captar todas as reuniões realizadas pelo conjunto de e-mails citados na lista, para analise da equipe de gestão


function importareventos() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DADOS");
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DASHBOARD");
  var calendarios = CalendarApp.getAllCalendars();
  
  const whiteList = [
    { email: "lucas.aguilar@wakke.co", id: 1 },
    { email: "marcelo.mattos@wakke.co", id: 2 },
    { email: "ana.rodrigues@wakke.co", id: 3 },
    { email: "caua.nicolau@wakke.co", id: 4 },
    { email: "ingrid.oliveira@wakke.co", id: 5 },
    { email: "vinicius.souza@wakke.co", id: 6 },
    { email: "vinicius.santiago@wakke.co", id: 7 },
    { email: "rodrigo.nunes@wakke.co", id: 8 },
    { email: "julia.lima@wakke.co", id: 9 },
    { email: "mariana.sartori@wakke.co", id: 10 },
    { email: "mariana.faustino@wakke.co", id: 11 },
    { email: "kethely.leal@wakke.co", id: 12 },
    { email: "maria.algislen@wakke.co", id: 13 },
    { email: "gabriel.conceicao@wakke.co", id: 14 }
  ];

  
const excluders = ["off","Horário de Café","DENTISTA", "SAÍDA DENTISTA", "Compromisso - Escola Alice", "Out of office", "almoço", "saida", "saída", "[off]", "ficar on no chat", "ficar off no chat","horário do café","horário de café" ,"café","Café"];

  
  const dtInicio = new Date(dashboard.getRange("B2").getValue());
  const dtFim = new Date(dashboard.getRange("C2").getValue());

  // Limpa os dados existentes na planilha antes de escrever novos dados
  planilha.clear();

  // Escrever dados do evento
  var dados = [];

  // Filtra os calendários que estão presentes na whitelist
  calendarios = calendarios.filter(f => whiteList.find(fw => fw.email.trim() == f.getId().trim()) != null );

  // Laço na lista de calendários
  calendarios.forEach((calendario, index) => {
    // Captura todos os eventos do calendário do ciclo que o título não esteja no array de excluders
    var evts = calendario.getEvents(dtInicio, dtFim).filter(f => excluders.indexOf(f.getTitle().trim().toLowerCase()) == -1 );

    console.log(`Calendário ${calendario.getName()} ${calendario.getId()} ${index + 1}: ${evts.length} `);
    
    evts.forEach(evento => {
      try {
        var title = evento.getTitle();
        var startTime = evento.getStartTime();
        var endTime = evento.getEndTime();
        var description = evento.getDescription();
        var criador = evento.getCreators().join(', '); // Assume que pode haver mais de um criador
        
        if (title && startTime && endTime) {
          // Inclui uma estrutura de dados 
          var duracaoMinutos = (endTime - startTime) / (60 * 1000); // Cálculo da duração em minutos
          dados.push([title, startTime, formatarHoraMinuto(startTime), formatarHoraMinuto(endTime), criador, whiteList.find(f => f.email == calendario.getId()).id, formatarHoraMinutoMinutos(duracaoMinutos)]);
        }
      } catch (error) {
        Logger.log("Erro ao processar evento: " + error);
        return;
      }
    });
  });

  // Escrever novos dados na planilha
  if (dados.length > 0) {
    // Ordena a lista por ID
    dados = dados.sort((a, b) => a[5] - b[5] );

    var linhaInicial = planilha.getLastRow() + 1;
    planilha.getRange(1, 1, 1, 7).setValues([["Descrição do evento","Data","Hora de início","Hora fim","Responsável","ID AGENDA","HORA/REUNIÃO"]]);
    planilha.getRange(linhaInicial + 1, 1, dados.length, dados[0].length).setValues(dados);
  }
}

// Função para formatar horas e minutos como "HH:mm"
function formatarHoraMinuto(data) {
  var hora = data.getHours();
  var minuto = data.getMinutes();

  var horaFormatada = hora < 10 ? "0" + hora : hora;
  var minutoFormatado = minuto < 10 ? "0" + minuto : minuto;

  return horaFormatada + ":" + minutoFormatado;
}

// Função para formatar minutos como "HH:mm"
function formatarHoraMinutoMinutos(minutos) {
  var horas = Math.floor(minutos / 60);
  var minutosRestantes = minutos % 60;

  var horaFormatada = horas < 10 ? "0" + horas : horas;
  var minutoFormatado = minutosRestantes < 10 ? "0" + minutosRestantes : minutosRestantes;

  return horaFormatada + ":" + minutoFormatado;
}
