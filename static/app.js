
document.addEventListener("DOMContentLoaded", () => {

    const form = document.getElementById("convert-form");
    const input = document.getElementById("file-input");
    const zone = document.getElementById("dropzone");

    const picked = document.getElementById("file-picked");
    const pickedName = document.getElementById("picked-name");
    const clear = document.getElementById("clear-file");

    const title = document.getElementById("upload-title");
    const types = document.getElementById("file-types");

    const button = form.querySelector(".convert-button");
    const buttonText = button.querySelector(".button-text");

    const MAX_SIZE = 16 * 1024 * 1024;


    /* =====================================================
       UPDATE SELECTED FILE
       ===================================================== */

    function updateFile(file) {

        if (!file) {

            picked.hidden = true;
            pickedName.textContent = "";

            return;
        }

        pickedName.textContent = file.name;

        picked.hidden = false;
    }


    /* =====================================================
       FILE VALIDATION
       ===================================================== */

    function validateFile(file) {

        if (!file) {
            return {
                valid: false,
                message: "Please select a file."
            };
        }


        /* Maximum file size */

        if (file.size > MAX_SIZE) {

            return {
                valid: false,
                message: "File size must be 16 MB or smaller."
            };
        }


        /* Current conversion mode */

        const selectedMode =
            document.querySelector(
                'input[name="mode"]:checked'
            );


        if (!selectedMode) {

            return {
                valid: false,
                message: "Please select a conversion type."
            };
        }


        const mode = selectedMode.value;

        const fileName =
            file.name.toLowerCase();


        /* Excel → PDF */

        if (mode === "excel-to-pdf") {

            if (
                !fileName.endsWith(".xlsx") &&
                !fileName.endsWith(".xlsm")
            ) {

                return {
                    valid: false,
                    message:
                        "Please select an XLSX or XLSM file."
                };
            }
        }


        /* PDF → Excel */

        if (mode === "pdf-to-excel") {

            if (!fileName.endsWith(".pdf")) {

                return {
                    valid: false,
                    message:
                        "Please select a PDF file."
                };
            }
        }


        return {
            valid: true
        };
    }


    /* =====================================================
       SHOW ERROR
       ===================================================== */

    function showError(message) {

        const oldError =
            document.querySelector(".client-error");

        if (oldError) {
            oldError.remove();
        }


        const error =
            document.createElement("div");

        error.className =
            "alert error client-error";

        error.setAttribute(
            "role",
            "alert"
        );

        error.innerHTML = `
            <span class="alert-icon">!</span>

            <div>
                <strong>Invalid file</strong>
                <p>${message}</p>
            </div>
        `;


        picked.insertAdjacentElement(
            "afterend",
            error
        );


        setTimeout(() => {

            error.style.opacity = "0";

            error.style.transform =
                "translateY(-5px)";

            error.style.transition =
                ".2s ease";


            setTimeout(() => {

                error.remove();

            }, 200);

        }, 3500);
    }


    /* =====================================================
       HANDLE FILE
       ===================================================== */

    function handleFile(file) {

        const result =
            validateFile(file);


        if (!result.valid) {

            input.value = "";

            updateFile(null);

            showError(result.message);

            return;
        }


        updateFile(file);
    }


    /* =====================================================
       FILE INPUT
       ===================================================== */

    input.addEventListener(
        "change",
        () => {

            const file =
                input.files[0];

            if (file) {
                handleFile(file);
            }

        }
    );


    /* =====================================================
       CLEAR FILE
       ===================================================== */

    clear.addEventListener(
        "click",
        () => {

            input.value = "";

            updateFile(null);

        }
    );


    /* =====================================================
       DRAG ENTER / DRAG OVER
       ===================================================== */

    ["dragenter", "dragover"].forEach(
        eventName => {

            zone.addEventListener(
                eventName,
                event => {

                    event.preventDefault();
                    event.stopPropagation();

                    zone.classList.add(
                        "dragover"
                    );

                }
            );

        }
    );


    /* =====================================================
       DRAG LEAVE / DROP
       ===================================================== */

    ["dragleave", "drop"].forEach(
        eventName => {

            zone.addEventListener(
                eventName,
                event => {

                    event.preventDefault();
                    event.stopPropagation();

                    zone.classList.remove(
                        "dragover"
                    );

                }
            );

        }
    );


    /* =====================================================
       DROP FILE
       ===================================================== */

    zone.addEventListener(
        "drop",
        event => {

            const files =
                event.dataTransfer.files;


            if (!files.length) {
                return;
            }


            const file =
                files[0];


            /*
             * Put the dropped file into
             * the actual input element.
             */

            try {

                const dataTransfer =
                    new DataTransfer();

                dataTransfer.items.add(file);

                input.files =
                    dataTransfer.files;

            } catch (error) {

                console.warn(
                    "Could not assign dropped file.",
                    error
                );

            }


            handleFile(file);
        }
    );


    /* =====================================================
       CONVERSION MODE
       ===================================================== */

    document
        .querySelectorAll('input[name="mode"]')
        .forEach(radio => {

            radio.addEventListener(
                "change",
                () => {

                    /*
                     * Update active card
                     */

                    document
                        .querySelectorAll(".mode")
                        .forEach(mode => {

                            const modeInput =
                                mode.querySelector(
                                    "input"
                                );

                            mode.classList.toggle(
                                "active",
                                modeInput.checked
                            );

                        });


                    const mode =
                        radio.value;


                    /*
                     * Excel → PDF
                     */

                    if (
                        mode ===
                        "excel-to-pdf"
                    ) {

                        input.accept =
                            ".xlsx,.xlsm";

                        title.textContent =
                            "Drop your Excel file here";

                        types.textContent =
                            "XLSX or XLSM · Maximum 16 MB";
                    }


                    /*
                     * PDF → Excel
                     */

                    else if (
                        mode ===
                        "pdf-to-excel"
                    ) {

                        input.accept =
                            ".pdf";

                        title.textContent =
                            "Drop your PDF here";

                        types.textContent =
                            "PDF · Maximum 16 MB";
                    }


                    /*
                     * Clear old file
                     * when mode changes.
                     */

                    input.value = "";

                    updateFile(null);

                }
            );

        });


    /* =====================================================
       FORM SUBMISSION
       ===================================================== */

    form.addEventListener(
        "submit",
        event => {

            const file =
                input.files[0];


            /*
             * Validate before sending.
             */

            const result =
                validateFile(file);


            if (!result.valid) {

                event.preventDefault();

                showError(
                    result.message
                );

                return;
            }


            /*
             * Prevent double submission.
             */

            button.disabled = true;


            /*
             * Activate loading state.
             */

            button.classList.add(
                "loading"
            );

            button.setAttribute(
                "aria-busy",
                "true"
            );


            /*
             * Change button text.
             */

            buttonText.textContent =
                "Converting...";


            /*
             * Keep arrow visible.
             */

            const arrow =
                button.querySelector(
                    ".button-arrow"
                );

            if (arrow) {
                arrow.textContent = "→";
            }

        }
    );


    /* =====================================================
       INITIAL STATE
       ===================================================== */

    const initialMode =
        document.querySelector(
            'input[name="mode"]:checked'
        );


    if (initialMode) {

        initialMode.dispatchEvent(
            new Event("change")
        );

    }

});

