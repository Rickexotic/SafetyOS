let touchStartX = 0;
let touchEndX = 0;

document.addEventListener(
    "touchstart",
    e => {

        touchStartX =
            e.changedTouches[0].screenX;
    }
);

document.addEventListener(
    "touchend",
    e => {

        touchEndX =
            e.changedTouches[0].screenX;

        handleSwipe();
    }
);

function handleSwipe() {

    const distance =
        touchEndX - touchStartX;

    if (Math.abs(distance) < 80) {
        return;
    }

    const page =
        window.location.pathname
            .split("/")
            .pop();

    /* REPORT */

    if (
        page === "index.html" ||
        page === ""
    ) {

        if (distance < -80) {

            window.location.href =
                "history.html";
        }
    }

    /* HISTORY */

    else if (
        page === "history.html"
    ) {

        if (distance > 80) {

            window.location.href =
                "index.html";
        }

        if (distance < -80) {

            window.location.href =
                "dashboard.html";
        }
    }

    /* DASHBOARD */

    else if (
        page === "dashboard.html"
    ) {

        if (distance > 80) {

            window.location.href =
                "history.html";
        }
    }
}
