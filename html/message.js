
var data_

function dataHandler(e){

    let lineType = document.getElementById('lineType')
    let lineSize = document.getElementById('lineSize')

    console.log('Parent received message!: ', e.data)
    if (e.data && e.data.lineSize && e.data.lineType ) {

            dataLineSize = e.data.lineSize
            dataLineType = e.data.lineType

    } 

    // if (e.data.lineSize == ''){
    //     dataLineSize = 'None'
    // }

    // if (e.data.lineType == ''){
    //     dataLineType = '100_point'
    // }
    
    lineType.value = dataLineType
    lineSize.value = dataLineSize

    return
}



var message = {
    type: 'iframe',
    commands: {}
}

var SendMessage = function() {
    parent.postMessage(message, '*')
}



var eventMethod = window.addEventListener ? "addEventListener" : "attachEvent"
var eventer = window[eventMethod]
var messageEvent = eventMethod == "attachEvent" ? "onmessage" : "message"

eventer(messageEvent, dataHandler)




