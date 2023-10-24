this.socket = new WebSocket('ws://localhost:5678');

this.socket.onmessage = (event) => {
    // console.log(event.data);
    eval(event.data)
};

this.socket.onclose = (event) => {
    if (event.wasClean) {
        console.log(`Connection closed cleanly, code=${event.code}, reason=${event.reason}`);
    } else {
        console.error('Connection died');
    }
};

this.socket.onerror = (error) => {
    console.error(`WebSocket Error: ${error}`);
};