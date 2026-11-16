(function () {
    let paginationRun = 0;

    function cloneNode(node) {
        return node.cloneNode(true);
    }

    function logPlacement(state, node, type) {
        const text = (node.textContent || "").replace(/\s+/g, " ").trim();
        const usedHeight = state.current.body.scrollHeight;
        const availableHeight = state.current.maxBodyHeight;
        console.log("[print-pagination] elemento inserido", {
            run: state.run,
            page: state.current.pageNumber,
            type: type,
            text: text,
            usedHeight: usedHeight,
            availableHeight: availableHeight,
            overflow: usedHeight > (availableHeight + 2),
        });
    }

    function createPage(headerTemplate, footerTemplate, pageClassName) {
        const page = document.createElement("section");
        page.className = pageClassName || "print-only print-sheet print-page";

        const header = cloneNode(headerTemplate);
        const body = document.createElement("div");
        body.className = "print-page-body";
        const footer = cloneNode(footerTemplate);

        page.appendChild(header);
        page.appendChild(body);
        page.appendChild(footer);

        return { page, header, body, footer, maxBodyHeight: 0 };
    }

    function refreshPageMetrics(pageRef) {
        const available = pageRef.page.clientHeight - pageRef.header.offsetHeight - pageRef.footer.offsetHeight;
        pageRef.maxBodyHeight = Math.max(0, available);
        pageRef.body.style.height = pageRef.maxBodyHeight + "px";
        pageRef.body.style.maxHeight = pageRef.maxBodyHeight + "px";
    }

    function createAndMountPage(state) {
        state.current = createPage(state.headerTemplate, state.footerTemplate, state.pageClassName);
        state.pageCount += 1;
        state.current.pageNumber = state.pageCount;
        // Paginas seguintes de uma folha herdam o subtitulo dela (ver print-sheet-break).
        if (state.sheetSubtitle) state.current.page.dataset.sheetSubtitle = state.sheetSubtitle;
        state.output.appendChild(state.current.page);
        refreshPageMetrics(state.current);
        console.log("[print-pagination] pagina criada", {
            run: state.run,
            page: state.current.pageNumber,
            availableHeight: state.current.maxBodyHeight,
        });
    }

    function fitsInCurrentPage(pageRef) {
        const tolerance = 2;
        const bottomReserve = 32;
        if (pageRef.body.scrollHeight > (pageRef.maxBodyHeight + tolerance)) return false;

        const bodyRect = pageRef.body.getBoundingClientRect();
        let contentBottom = 0;
        const nodes = Array.from(pageRef.body.querySelectorAll("*"));

        nodes.forEach((node) => {
            const style = window.getComputedStyle(node);
            if (node.tagName === "STYLE" || node.tagName === "SCRIPT" || style.position === "fixed" || style.display === "none") return;

            // A margem inferior nao entra na conta: espaco vazio no fim da pagina nao ocupa area impressa.
            const rect = node.getBoundingClientRect();
            contentBottom = Math.max(contentBottom, rect.bottom - bodyRect.top);
        });

        return contentBottom <= (pageRef.body.clientHeight - bottomReserve);
    }

    function applyPageCountersAndState(output) {
        const pages = Array.from(output.querySelectorAll(".print-page"));
        const total = pages.length;
        // Paginas marcadas com data-sheet-start iniciam uma nova folha (ex.: folha de conteudo do romaneio):
        // "continuação" e "continua na próxima página" só valem entre páginas da mesma folha.
        const startsSheet = (page) => page.dataset.sheetStart === "1";
        // Cada folha tem numeracao propria ("Pág. 1 de 2" ... "Pág. 1 de 3"); sem folhas, e o relatorio inteiro.
        const sheetNumbers = [];
        let sheetFirstIndex = 0;
        pages.forEach((page, index) => {
            if (index > 0 && startsSheet(page)) sheetFirstIndex = index;
            sheetNumbers.push({ first: sheetFirstIndex, current: index - sheetFirstIndex + 1, total: 0 });
        });
        sheetNumbers.forEach((entry, index) => {
            const isLastOfSheet = index === total - 1 || sheetNumbers[index + 1].first !== entry.first;
            if (!isLastOfSheet) return;
            for (let i = entry.first; i <= index; i += 1) sheetNumbers[i].total = entry.current;
        });
        pages.forEach((page, index) => {
            const current = index + 1;
            const continuesFromPrevious = index > 0 && !startsSheet(page);
            const continuesOnNext = current < total && !startsSheet(pages[index + 1]);
            page.querySelectorAll(".print-page-current").forEach((el) => {
                el.textContent = String(sheetNumbers[index].current);
            });
            page.querySelectorAll(".print-page-total").forEach((el) => {
                el.textContent = String(sheetNumbers[index].total);
            });
            if (page.dataset.sheetSubtitle) {
                page.querySelectorAll(".print-subtitle").forEach((el) => {
                    el.textContent = page.dataset.sheetSubtitle;
                });
            }
            if (continuesFromPrevious) {
                page.querySelectorAll(".print-title").forEach((el) => {
                    if (!el.textContent.includes("(continuação)")) {
                        el.textContent = el.textContent + " (continuação)";
                    }
                });
            }
            const state = page.querySelector(".print-footer-state");
            if (state) {
                if (continuesOnNext) state.textContent = "CONTINUA NA PRÓXIMA PÁGINA";
                else state.textContent = current === total ? "FIM DO RELATÓRIO" : "";
            }
        });
    }

    function appendBlockWithOptionalTitle(state, blockNode, titleNode) {
        const titleClone = titleNode ? cloneNode(titleNode) : null;
        const blockClone = cloneNode(blockNode);

        if (titleClone) state.current.body.appendChild(titleClone);
        state.current.body.appendChild(blockClone);

        if (fitsInCurrentPage(state.current)) return;

        if (titleClone) titleClone.remove();
        blockClone.remove();

        createAndMountPage(state);

        if (titleClone) state.current.body.appendChild(titleClone);
        state.current.body.appendChild(blockClone);
    }

    function appendTreeSplitByItems(state, treeNode, sectionTitleNode) {
        const items = Array.from(treeNode.querySelectorAll(":scope > .print-tree-item"));
        if (!items.length) {
            appendBlockWithOptionalTitle(state, treeNode, sectionTitleNode);
            return;
        }

        let treeContainer = treeNode.cloneNode(false);

        function startNewPageWithTitle(asContinuation) {
            createAndMountPage(state);

            if (sectionTitleNode) {
                const repeatedTitle = cloneNode(sectionTitleNode);
                if (asContinuation) {
                    repeatedTitle.textContent = repeatedTitle.textContent + " (continuação)";
                }
                state.current.body.appendChild(repeatedTitle);
            }

            treeContainer = treeNode.cloneNode(false);
            state.current.body.appendChild(treeContainer);
        }

        if (sectionTitleNode) {
            state.current.body.appendChild(cloneNode(sectionTitleNode));
        }
        state.current.body.appendChild(treeContainer);

        items.forEach((item) => {
            const itemClone = cloneNode(item);
            treeContainer.appendChild(itemClone);

            if (fitsInCurrentPage(state.current)) return;

            itemClone.remove();

            const treeIsEmpty = treeContainer.children.length === 0;
            if (treeIsEmpty) {
                treeContainer.appendChild(itemClone);
                return;
            }

            startNewPageWithTitle(true);
            treeContainer.appendChild(itemClone);
        });
    }

    function appendRecordCardsSplitByItems(state, listNode, sectionTitleNode, continuationHeaderNode, skipInitialScaffold) {
        const listChildren = Array.from(listNode.children);
        const items = listChildren.filter((child) => child.classList.contains("print-record-card"));
        const firstItemIndex = listChildren.findIndex((child) => child.classList.contains("print-record-card"));
        const embeddedScaffoldNodes = firstItemIndex > 0 ? listChildren.slice(0, firstItemIndex) : [];
        const scaffoldNodes = embeddedScaffoldNodes.length ? embeddedScaffoldNodes : (continuationHeaderNode ? [continuationHeaderNode] : []);
        if (!items.length) {
            if (continuationHeaderNode) {
                appendBlockWithOptionalTitle(state, continuationHeaderNode, sectionTitleNode);
                appendBlockWithOptionalTitle(state, listNode, null);
                return;
            }
            appendBlockWithOptionalTitle(state, listNode, sectionTitleNode);
            return;
        }

        let listContainer;
        let currentTitleClone = null;
        let pageHadContentBeforeList = false;

        function mountList(asContinuation) {
            currentTitleClone = null;
            pageHadContentBeforeList = state.current.body.children.length > 0;

            if (sectionTitleNode && (!skipInitialScaffold || asContinuation)) {
                const repeatedTitle = cloneNode(sectionTitleNode);
                if (asContinuation) {
                    repeatedTitle.textContent = repeatedTitle.textContent + " - CONTINUA\u00c7\u00c3O";
                }
                currentTitleClone = repeatedTitle;
                state.current.body.appendChild(repeatedTitle);
            }

            listContainer = listNode.cloneNode(false);
            if (!skipInitialScaffold || asContinuation) {
                scaffoldNodes.forEach((node) => listContainer.appendChild(cloneNode(node)));
            }
            state.current.body.appendChild(listContainer);
        }

        function startNewPageWithTitle(asContinuation) {
            createAndMountPage(state);
            mountList(asContinuation);
        }

        function removeCurrentList() {
            if (currentTitleClone) currentTitleClone.remove();
            listContainer.remove();
        }

        mountList(false);
        if (!fitsInCurrentPage(state.current) && pageHadContentBeforeList) {
            removeCurrentList();
            startNewPageWithTitle(false);
        }

        items.forEach((item) => {
            const itemClone = cloneNode(item);
            listContainer.appendChild(itemClone);

            if (fitsInCurrentPage(state.current)) {
                logPlacement(state, itemClone, "item-lista");
                return;
            }

            itemClone.remove();

            const listIsEmpty = !listContainer.querySelector(":scope > .print-record-card");
            if (listIsEmpty) {
                if (pageHadContentBeforeList) {
                    removeCurrentList();
                    startNewPageWithTitle(false);
                    listContainer.appendChild(itemClone);
                    logPlacement(state, itemClone, "item-lista");
                    return;
                }
                listContainer.appendChild(itemClone);
                logPlacement(state, itemClone, "item-lista");
                return;
            }

            startNewPageWithTitle(true);
            listContainer.appendChild(itemClone);
            logPlacement(state, itemClone, "item-lista");
        });
    }

    function cloneTableShell(tableNode, head, body) {
        const tableContainer = tableNode.cloneNode(false);
        let bodyContainer = null;

        Array.from(tableNode.children).forEach((child) => {
            if (child === head) {
                tableContainer.appendChild(cloneNode(head));
                return;
            }

            if (child === body) {
                bodyContainer = body.cloneNode(false);
                tableContainer.appendChild(bodyContainer);
                return;
            }

            tableContainer.appendChild(cloneNode(child));
        });

        if (!bodyContainer) {
            bodyContainer = body.cloneNode(false);
            tableContainer.appendChild(bodyContainer);
        }

        return { tableContainer, bodyContainer };
    }

    function appendTableSplitByRows(state, tableNode, sectionTitleNode) {
        const head = tableNode.querySelector(":scope > thead");
        const body = tableNode.querySelector(":scope > tbody");
        const rows = body ? Array.from(body.querySelectorAll(":scope > tr")) : [];
        if (!head || !body || !rows.length) {
            appendBlockWithOptionalTitle(state, tableNode, sectionTitleNode);
            return;
        }

        let tableContainer;
        let bodyContainer;
        let currentTitleClone = null;
        let pageHadContentBeforeTable = false;

        function mountTable(asContinuation) {
            currentTitleClone = null;
            pageHadContentBeforeTable = state.current.body.children.length > 0;
            if (sectionTitleNode) {
                const repeatedTitle = cloneNode(sectionTitleNode);
                if (asContinuation) {
                    repeatedTitle.textContent = repeatedTitle.textContent + " (continuação)";
                }
                currentTitleClone = repeatedTitle;
                state.current.body.appendChild(repeatedTitle);
            }

            const tableShell = cloneTableShell(tableNode, head, body);
            tableContainer = tableShell.tableContainer;
            bodyContainer = tableShell.bodyContainer;
            state.current.body.appendChild(tableContainer);
        }

        function startNewPageWithTitle(asContinuation) {
            createAndMountPage(state);
            mountTable(asContinuation);
        }

        function removeCurrentTable() {
            if (currentTitleClone) currentTitleClone.remove();
            tableContainer.remove();
        }

        mountTable(false);
        if (!fitsInCurrentPage(state.current) && pageHadContentBeforeTable) {
            removeCurrentTable();
            startNewPageWithTitle(false);
        }

        rows.forEach((row) => {
            const rowClone = cloneNode(row);
            bodyContainer.appendChild(rowClone);

            if (fitsInCurrentPage(state.current)) {
                logPlacement(state, rowClone, "linha-tabela");
                return;
            }

            rowClone.remove();

            const tableIsEmpty = bodyContainer.children.length === 0;
            if (tableIsEmpty) {
                if (fitsInCurrentPage(state.current) && pageHadContentBeforeTable) {
                    if (currentTitleClone) currentTitleClone.remove();
                    tableContainer.remove();
                    startNewPageWithTitle(false);
                    bodyContainer.appendChild(rowClone);
                    logPlacement(state, rowClone, "linha-tabela");
                    return;
                }
                if (currentTitleClone) currentTitleClone.remove();
                tableContainer.remove();
                startNewPageWithTitle(false);
                bodyContainer.appendChild(rowClone);
                logPlacement(state, rowClone, "linha-tabela");
                return;
            }

            startNewPageWithTitle(true);
            bodyContainer.appendChild(rowClone);
            logPlacement(state, rowClone, "linha-tabela");
        });
    }

    function getDirectSplittableTable(containerNode) {
        if (!containerNode.children) return null;

        const tables = Array.from(containerNode.children).filter(
            (child) => child.classList && child.classList.contains("print-table")
        );
        if (tables.length !== 1 || tables[0].nextElementSibling) return null;

        return tables[0];
    }

    function isRecordListHead(node) {
        return node.classList.contains("print-dashboard-list-head") || node.classList.contains("print-record-list-head");
    }

    function appendContainerTableSplitByRows(state, containerNode, tableNode, sectionTitleNode) {
        const head = tableNode.querySelector(":scope > thead");
        const body = tableNode.querySelector(":scope > tbody");
        const rows = body ? Array.from(body.querySelectorAll(":scope > tr")) : [];
        if (!head || !body || !rows.length) {
            appendBlockWithOptionalTitle(state, containerNode, sectionTitleNode);
            return;
        }

        const prefixNodes = [];
        for (const child of Array.from(containerNode.children)) {
            if (child === tableNode) break;
            prefixNodes.push(child);
        }

        let containerClone;
        let bodyContainer;
        let currentTitleClone = null;
        let pageHadContentBeforeContainer = false;

        function mountContainer(asContinuation) {
            currentTitleClone = null;
            pageHadContentBeforeContainer = state.current.body.children.length > 0;
            if (sectionTitleNode) {
                const repeatedTitle = cloneNode(sectionTitleNode);
                if (asContinuation) {
                    repeatedTitle.textContent = repeatedTitle.textContent + " (continuação)";
                }
                currentTitleClone = repeatedTitle;
                state.current.body.appendChild(repeatedTitle);
            }

            containerClone = containerNode.cloneNode(false);
            prefixNodes.forEach((node) => {
                const prefixClone = cloneNode(node);
                if (asContinuation && prefixClone.classList.contains("print-section-title")) {
                    prefixClone.textContent = prefixClone.textContent + " (continuação)";
                }
                containerClone.appendChild(prefixClone);
            });

            const tableShell = cloneTableShell(tableNode, head, body);
            bodyContainer = tableShell.bodyContainer;
            containerClone.appendChild(tableShell.tableContainer);
            state.current.body.appendChild(containerClone);
        }

        function startNewPageWithTitle(asContinuation) {
            createAndMountPage(state);
            mountContainer(asContinuation);
        }

        function removeCurrentContainer() {
            if (currentTitleClone) currentTitleClone.remove();
            containerClone.remove();
        }

        mountContainer(false);
        if (!fitsInCurrentPage(state.current) && pageHadContentBeforeContainer) {
            removeCurrentContainer();
            startNewPageWithTitle(false);
        }

        rows.forEach((row) => {
            const rowClone = cloneNode(row);
            bodyContainer.appendChild(rowClone);

            if (fitsInCurrentPage(state.current)) {
                logPlacement(state, rowClone, "linha-tabela-aninhada");
                return;
            }

            rowClone.remove();

            const tableIsEmpty = bodyContainer.children.length === 0;
            if (tableIsEmpty) {
                removeCurrentContainer();
                startNewPageWithTitle(false);
                bodyContainer.appendChild(rowClone);
                logPlacement(state, rowClone, "linha-tabela-aninhada");
                return;
            }

            startNewPageWithTitle(true);
            bodyContainer.appendChild(rowClone);
            logPlacement(state, rowClone, "linha-tabela-aninhada");
        });
    }

    function prepareReport(report) {
        const headerTemplate = report.querySelector(":scope > .print-header");
        const footerTemplate = report.querySelector(":scope > .print-footer");
        if (!headerTemplate || !footerTemplate) return;

        let output = report.nextElementSibling;
        if (!output || !output.classList.contains("print-pages-output")) {
            output = document.createElement("div");
            output.className = "print-only print-pages-output";
            report.insertAdjacentElement("afterend", output);
        }
        output.innerHTML = "";

        const nodes = Array.from(report.children).filter(
            (node) => !node.classList.contains("print-header") && !node.classList.contains("print-footer")
        );

        const state = {
            headerTemplate: headerTemplate,
            footerTemplate: footerTemplate,
            output: output,
            current: null,
            pageCount: 0,
            run: paginationRun,
            pageClassName: Array.from(report.classList)
                .filter((className) => className !== "js-print-report")
                .concat("print-page")
                .join(" "),
        };

        createAndMountPage(state);

        let pendingSectionTitle = null;
        let pendingContinuationHeader = null;
        nodes.forEach((node) => {
            if (node.classList.contains("print-page-break-before")) {
                if (state.current.body.children.length) createAndMountPage(state);
                if (node.classList.contains("print-sheet-break")) {
                    state.current.page.dataset.sheetStart = "1";
                    if (node.dataset.sheetSubtitle) {
                        state.sheetSubtitle = node.dataset.sheetSubtitle;
                        state.current.page.dataset.sheetSubtitle = state.sheetSubtitle;
                    }
                }
                if (node.classList.contains("print-page-break-marker")) return;
            }
            if (node.classList.contains("print-section-title")) {
                pendingSectionTitle = node;
                return;
            }

            if (isRecordListHead(node)) {
                const nextElement = node.nextElementSibling;
                if (nextElement && nextElement.classList.contains("print-record-list")) {
                    pendingContinuationHeader = node;
                    return;
                }
            }

            if (node.classList.contains("print-tree")) {
                appendTreeSplitByItems(state, node, pendingSectionTitle);
                pendingSectionTitle = null;
                return;
            }

            if (node.classList.contains("print-record-list")) {
                const continuationHeaderNode = pendingContinuationHeader;
                pendingContinuationHeader = null;
                appendRecordCardsSplitByItems(
                    state,
                    node,
                    pendingSectionTitle,
                    continuationHeaderNode,
                    false
                );
                pendingSectionTitle = null;
                return;
            }

            if (node.classList.contains("print-table")) {
                appendTableSplitByRows(state, node, pendingSectionTitle);
                pendingSectionTitle = null;
                return;
            }

            const directSplittableTable = getDirectSplittableTable(node);
            if (directSplittableTable) {
                appendContainerTableSplitByRows(state, node, directSplittableTable, pendingSectionTitle);
                pendingSectionTitle = null;
                return;
            }

            appendBlockWithOptionalTitle(state, node, pendingSectionTitle);
            pendingSectionTitle = null;
        });

        applyPageCountersAndState(output);
    }

    function prepareAllReports() {
        paginationRun += 1;
        console.log("[print-pagination] inicio", { run: paginationRun });
        const reports = Array.from(document.querySelectorAll(".js-print-report"));
        document.body.classList.add("print-pagination-measure");
        reports.forEach((report) => prepareReport(report));
        document.body.classList.remove("print-pagination-measure");
        if (reports.length) {
            document.body.classList.add("print-pagination-ready");
        }
    }

    window.prepareStandardPrint = prepareAllReports;

    window.runStandardPrint = function () {
        prepareAllReports();
        window.setTimeout(function () {
            window.print();
        }, 30);
    };

    window.addEventListener("beforeprint", prepareAllReports);
})();
