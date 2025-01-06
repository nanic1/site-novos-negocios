// Gráfico da Origem
am5.ready(function() {

    // Create root element
    var root = am5.Root.new("origemG");

    // Set themes
    root.setThemes([
      am5themes_Animated.new(root)
    ]);

    // Create chart
    var chart = root.container.children.push(am5percent.PieChart.new(root, {
      layout: root.verticalLayout,
      innerRadius: am5.percent(70)
    }));

    // Create series
    var series = chart.series.push(am5percent.PieSeries.new(root, {
      valueField: "value",
      categoryField: "category",
      alignLabels: false
    }));

    // Remove the default label format (percent) and set it to display value
    series.labels.template.setAll({
      text: "{value}",  // Exibe o valor absoluto ao invés da porcentagem
      populateText: true
    });

    // Set data
    series.data.setAll([
      { value: 34, category: "Externo" },
      { value: 35, category: "Interno" }
    ]);

    // Play initial series animation
    series.appear(1000, 100);

}); // end am5.ready()


am5.ready(function() {

  // Create root element
  // https://www.amcharts.com/docs/v5/getting-started/#Root_element
  var root = am5.Root.new("agenteG");
  
  
  // Set themes
  // https://www.amcharts.com/docs/v5/concepts/themes/
  root.setThemes([
    am5themes_Animated.new(root)
  ]);
  
  
  // Create chart
  // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/
  var chart = root.container.children.push(am5percent.PieChart.new(root, {
    layout: root.verticalLayout
  }));
  
  
  // Create series
  // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Series
  var series = chart.series.push(am5percent.PieSeries.new(root, {
    valueField: "value",
    categoryField: "category"
  }));
  
  series.labels.template.setAll({
    text: "{value}" // Exibe apenas o valor
  });

  series.slices.template.setAll({
    tooltipText: "{category}: {value}"
  });
  
  // Set data
  // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Setting_data
  series.data.setAll([
    { value: 10, category: "PRISCYLLA" },
    { value: 9, category: "ALX" },
    { value: 6, category: "ALEX" },
    { value: 5, category: "MARCOS SP" },
    { value: 4, category: "ARTUR" },
    { value: 3, category: "MARCOS" },
  ]);
  
  
  // Create legend
  // https://www.amcharts.com/docs/v5/charts/percent-charts/legend-percent-series/
  var legend = chart.children.push(am5.Legend.new(root, {
    centerX: am5.percent(50),
    x: am5.percent(50),
    marginTop: 15,
    marginBottom: 15
  }));
  
  legend.data.setAll(series.dataItems);

  legend.valueLabels.template.setAll({
    forceHidden: true
  });
  
  // Play initial series animation
  // https://www.amcharts.com/docs/v5/concepts/animations/#Animation_of_series
  series.appear(1000, 100);
  
  }); // end am5.ready()


  // STATUS
  am5.ready(function() {

    // Create root element
    // https://www.amcharts.com/docs/v5/getting-started/#Root_element
    var root = am5.Root.new("statusG");
    
    
    // Set themes
    // https://www.amcharts.com/docs/v5/concepts/themes/
    root.setThemes([
      am5themes_Animated.new(root)
    ]);
    
    
    // Create chart
    // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/
    var chart = root.container.children.push(am5percent.PieChart.new(root, {
      layout: root.verticalLayout
    }));
    
    
    // Create series
    // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Series
    var series = chart.series.push(am5percent.PieSeries.new(root, {
      valueField: "value",
      categoryField: "category"
    }));
  
    series.slices.template.setAll({
      tooltipText: "{category}: {value}"
    });

    series.labels.template.setAll({
      text: "{value}" // Exibe apenas o valor
    });
    
    // Set data
    // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Setting_data
    series.data.setAll([
      { value: 10, category: "CADASTRO" },
      { value: 9, category: "CALL" },
      { value: 6, category: "COMITE" },
      { value: 5, category: "CONTRATO" },
      { value: 4, category: "POS CALL" },
      { value: 3, category: "PRE ANALISE" },
    ]);
    
    
    // Create legend
    // https://www.amcharts.com/docs/v5/charts/percent-charts/legend-percent-series/
    var legend = chart.children.push(am5.Legend.new(root, {
      centerX: am5.percent(50),
      x: am5.percent(50),
      marginTop: 15,
      marginBottom: 15
    }));
    
    legend.data.setAll(series.dataItems);
  
    legend.labels.template.setAll({
      fontSize: 14 // Change the font size as needed
    });
    
    legend.valueLabels.template.setAll({
      forceHidden: true
    });

    // Play initial series animation
    // https://www.amcharts.com/docs/v5/concepts/animations/#Animation_of_series
    series.appear(1000, 100);
    
    }); // end am5.ready()


    // CONTAGEM
    am5.ready(function() {

      // Create root element
      // https://www.amcharts.com/docs/v5/getting-started/#Root_element
      var root = am5.Root.new("contagemG");
      
      
      // Set themes
      // https://www.amcharts.com/docs/v5/concepts/themes/
      root.setThemes([
        am5themes_Animated.new(root)
      ]);
      
      
      // Create chart
      // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/
      var chart = root.container.children.push(am5percent.PieChart.new(root, {
        layout: root.verticalLayout
      }));
      
      
      // Create series
      // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Series
      var series = chart.series.push(am5percent.PieSeries.new(root, {
        valueField: "value",
        categoryField: "category"
      }));
    
      series.slices.template.setAll({
        tooltipText: "{category}: {value}"
      });
  
      series.labels.template.setAll({
        text: "{value}" // Exibe apenas o valor
      });
      
      // Set data
      // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Setting_data
      series.data.setAll([
        { value: 10, category: "PRIMEIRA OP." },
        { value: 9, category: "FORMALIZACAO DO CONT." },
        { value: 6, category: "COMITE" },
        { value: 5, category: "VISITA" },
        { value: 4, category: "CADASTRO" },
        { value: 3, category: "CONTRATO" },
        { value: 7, category: "ENVIO DE DOC."},
        { value: 1, category: "POS-CALL"},
        { value: 5, category: "ENVIO DE PROPOSTA"},
      ]);
      
      
      // Create legend
      // https://www.amcharts.com/docs/v5/charts/percent-charts/legend-percent-series/
      var legend = chart.children.push(am5.Legend.new(root, {
        centerX: am5.percent(50),
        x: am5.percent(50),
        marginTop: 15,
        marginBottom: 15
      }));
      
      legend.data.setAll(series.dataItems);
    
      legend.labels.template.setAll({
        fontSize: 14 // Change the font size as needed
      });
      
      legend.valueLabels.template.setAll({
        forceHidden: true
      });
  
      // Play initial series animation
      // https://www.amcharts.com/docs/v5/concepts/animations/#Animation_of_series
      series.appear(1000, 100);
      
      }); // end am5.ready()

const XLSX = require("xlsx")

const arquivo = XLSX.readFile("BaseDados.xlsx"); // Selecionando arquivo de trabalho Excel

const selectPlanilha = arquivo.SheetNames[0] // Selecionando a planila; 0 = Folha1
const planilha = arquivo.Sheets[selectPlanilha] // Puxando todos os dados da planilha.

const enderecoExterno = "B2"; // Endereco da celula Externo
const enderecoInterno = "B3"; // Endereco da celula Interno
const celulaExterno = planilha[enderecoExterno];
const celulaInterno = planilha[enderecoInterno];

var valorExterno = celulaExterno ? celulaExterno.v : null; // Valor em INT Externo
var valorInterno = celulaInterno ? celulaInterno.v : null; // Valor em INT Interno