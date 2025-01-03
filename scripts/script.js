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
    var chart = root.container.children.push(
      am5percent.PieChart.new(root, {
        endAngle: 270
      })
    );
    
    // Create series
    // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Series
    var series = chart.series.push(
      am5percent.PieSeries.new(root, {
        valueField: "value",
        categoryField: "category",
        endAngle: 270
      })
    );
    
    series.states.create("hidden", {
      endAngle: -90
    });
    
    // Set data
    // https://www.amcharts.com/docs/v5/charts/percent-charts/pie-chart/#Setting_data
    series.data.setAll([{
      category: "Priscylla",
      value: 22
    }, {
      category: "ALX",
      value: 6
    }, {
      category: "Alex",
      value: 5
    }, {
      category: "Captação inicial",
      value: 5
    }, {
      category: "Marcos SP",
      value: 5
    }, {
      category: "Artur",
      value: 128.3
    }, {
      category: "UK",
      value: 99
    }]);
    
    series.appear(1000, 100);
    
    }); // end am5.ready()