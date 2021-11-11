# Exerc-cio-2
nome : lista 3 - 2
descrição : > -
  Layout do suplemento: só um botão com o texto "Formatar Planilha".
  Comportamento do suplemento: o script deve começar a remover qualquer
  formatação prévia nas células utilizadas na planilha. Em seguida, deve
  formatar uma planilha de notas de alunos previamente aberta no Excel (ou seja,
  use uma planilha sem formatação gerada como resultado do exercício anterior),
  aplicando uma formatação do tipo tabela do Excel com o "Estilo Médio 18".
  Depois da formatação como tabela, ordene os nomes dos alunos alfabeticamente,
  aplique formatação de número com uma casa decimal em todas as colunas com
  valores numéricos, e pinte em vermelho o texto das médias finais inferiores a
  6
  Tempo estimado para conclusão: 1 hora.
anfitrião : EXCEL
api_set : {}
script :
  conteúdo : >
    run.addEventListener ("click", async () => {
      esperar Excel.run (função assíncrona (contexto) {
        let sheet = context.workbook.worksheets.getActiveWorksheet ();
        let table = sheet.tables.getItemOrNullObject ("TabelaAluno");
        table.load ("isNullObject");
        aguarde context.sync ();
        if (table.isNullObject) {
          alert ("Atenção: Executar o suplemento do exercício` Lista 3 - 1`, \ n antes de executar este! ");
        }
        // altera o estilo
        table.style = 'TableStyleMedium18';
        table.load ('tableStyle');
        // ordena o nome dos alunos
        let sortFields = [
          {
            chave: 0, // Coluna nome do aluno
            ascendente: verdadeiro,
          }
        ];
        table.sort.apply (sortFields);
        // formata numeros
        formatos const = [
          ["0,00"]
        ];
        let body = table.getDataBodyRange ();
        var range = body.getColumnsAfter (-3);
        range.numberFormat = Formatos;
        // pinta as medias abaixo de 6 de vermelho
        let column = body.getLastColumn ();
        column.load ("rowCount");
        aguarde context.sync ();
        para (deixe i = 0; i <column.rowCount; i ++) {
          deixe row = column.getRow (i);
          row.load ("valores");
          aguarde context.sync ();
          if (row.values ​​[0] [0] <6) {
            row.format.font.color = "Vermelho";
          };
        };
        // preenchimento automático
        table.getRange (). format.autofitColumns ();
        table.getRange (). format.autofitRows ();
        return context.sync ();
      }). catch (função (erro) {
        console.log ("Erro:" + erro);
        if (instância de erro de OfficeExtension.Error) {
          console.log ("Informações de depuração:" + JSON.stringify (error.debugInfo));
        }
      });
    });
    // Força somente números nos inputs
    //
    https://www.geeksforgeeks.org/how-to-force-input-field-to-enter-numbers-only-using-javascript/
    function onlyNumberKey (evt) {
      // Somente caracteres ASCII permitidos neste intervalo
      var ASCIICode = evt.which? evt.which: evt.keyCode;
      if (ASCIICode> 31 && (ASCIICode <48 || ASCIICode> 57)) return false;
      return true;
    }
  linguagem : texto datilografado
modelo :
  conteúdo : | -
    <button id = "run" class = "ms-Button">
        <span class = "ms-Button-label"> Formatar </span>
    </button>
  linguagem : html
estilo :
  conteúdo : | -
    section.samples {
        margem superior: 20px;
    }
    section.samples .ms-Button, section.setup .ms-Button {
        display: bloco;
        margin-bottom: 5px;
        margem esquerda: 20px;
        largura mínima: 80px;
    }
  idioma : css
bibliotecas : |
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @ types / office-js
  office-ui-fabric-js@1.4.0/dist/css/fabric.min.css
  office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css
  core-js@2.4.1/client/core.min.js
  @ types / core-js
  jquery@3.1.1
  @ types / jquery @ 3.3.1
