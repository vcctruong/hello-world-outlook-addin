Office.onReady(function () {
    document.getElementById("helloButton").onclick = sayHello;
});

function sayHello() {
    const messageElement = document.getElementById("message");
    messageElement.textContent = "Hello World! This is your Outlook add-in.";
}
