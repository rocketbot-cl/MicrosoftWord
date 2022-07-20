let lineType = document.getElementById('lineType')
let lineSize = document.getElementById('lineSize')
let btnSelect = document.getElementById('btnSelect')



let styles = ["Dash dot", 
               "Dash and two dots", 
               "Dash dot stroked", 
               "Dash with large gap", 
               "Dash with small gap", 
               "Dotted", 
               "Double", 
               "Double wavy", 
               "3D Embossed look", 
               "3D Engraved look", 
               "Inset", 
               "None", 
               "Outset", 
               "Single", 
               "Single wavy", 
               "Thick thin large gap", 
               "Thick thin med gap", 
               "Thick thin small gap", 
               "Thin thick large gap", 
               "Thin thick med gap", 
               "Thin thick small gap", 
               "Thin thick thin large gap", 
               "Thin thick thin med gap", 
               "Thin thick thin small gap", 
               "Triple"]

let sizes = ["1/4 point", 
               "1/2 point", 
               "3/4 point", 
               "1 point", 
               "1 1/2 points", 
               "2 1/4 points", 
               "3 points", 
               "4 1/2 points", 
               "6 points"]

let valuesSizes = ["25_point",
                   "50_point", 
                   "75_point", 
                   "100_point", 
                   "150_points", 
                   "225_points", 
                   "300_points", 
                   "450_points", 
                   "600_points"]




function showSizes(array, arrayValues)  {

    while (lineSize.length > 0) {
        lineSize.remove(0);
    }

    for(item in array){
        let option = document.createElement('option')
        option.value = arrayValues[item]
        option.text = array[item]
        lineSize.appendChild(option)
    } 
}

lineType.addEventListener('change', function(){
    let valor = lineType.value
    let cut = []
    let cutSizes = []

    switch (valor) {
        case 'thin_thick_thin_med_gap':
            cut = sizes.slice(0, 8)
            cutSizes = valuesSizes.slice(0, 8)
            showSizes(cut, cutSizes, lineSize)
        break;
        case 'single_wavy':
            cut = ["3/4 point", "1 1/2 points"]
            cutSizes = ["75_point", "150_points"]
            showSizes(cut, cutSizes, lineSize)

        break;
        case 'double_wavy':
            cut = ["3/4 point"]
            cutSizes = ["75_point"]
            showSizes(cut, cutSizes, lineSize)
        break;
        case 'dash_dot_stroked':
            cut = ["3 points"]
            cutSizes = ["300_points"]
            showSizes(cut, cutSizes, lineSize)
        break;
        case 'double':
            cut = sizes.slice(0, 7)
            cutSizes = valuesSizes.slice(0, 7)
            showSizes(cut, cutSizes, lineSize)
        break;
        case 'triple':
            cut = sizes.slice(0, 7)
            cutSizes = valuesSizes.slice(0, 7)
            showSizes(cut, cutSizes, lineSize)
        break;
        default:
            showSizes(sizes, valuesSizes, lineSize)
            break;
    }

    message.commands["lineType"] = lineType.value
    message.commands["lineSize"] = lineSize.value
    SendMessage()
})

lineSize.addEventListener('change', function(){
    message.commands["lineType"] = lineType.value
    message.commands["lineSize"] = lineSize.value
    SendMessage()
})

// btnSelect.addEventListener('click', function(){
//     message.commands["lineType"] = lineType.value
//     message.commands["lineSize"] = lineSize.value
//     SendMessage()
// })

