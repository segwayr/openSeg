function startCalc(){
    // 入力された数値を元に計算
    var avr = document.tool.firstHit.value / document.tool.rTotal.value * document.tool.roundAVG.value;
    var worth = ((avr / document.tool.firstHit.value) - (250 / document.tool.spningRate.value)) * (100 / document.tool.enob.value);
    var func01 = document.tool.tnol.value * worth * document.tool.eRatio.value / 100;
    var solid = (avr / document.tool.firstHit.value * (100 / document.tool.enob.value)) - (1000 / document.tool.spningRate.value);
    var func02 = document.tool.tnol.value * solid * (100 - document.tool.eRatio.value) / 100;
    var totalProfit = func01 + func02;
    var perHour = totalProfit / document.tool.operatingTime.value;
    var eRatioOther = "" + document.tool.eRatio.value   
    totalProfit = totalProfit.toFixed(0);
    perHour = perHour.toFixed(0);
    if (document.tool.enob.value == "25"){
        var eRatioOther = "";
    }
    // 結果出力
    form.result.value="Expected Value"+totalProfit+"yen\rHourly pay"+perHour+"yen\r\r"+document.tool.spningRate.value+" "+document.tool.roundAVG.value+" "+eRatioOther+"Hourly pay"+perHour+"yen";
}
