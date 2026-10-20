document.addEventListener("DOMContentLoaded", () => {
    // --- ESTADO GLOBAL ---
    let siteConfigsCache = {};
    let activeSiteId = "default";
    let selectedFile = null;

    // --- NAVEGAÇÃO POR ABAS ---
    const navButtons = document.querySelectorAll(".nav-btn");
    const tabContents = document.querySelectorAll(".tab-content");

    navButtons.forEach(btn => {
        btn.addEventListener("click", () => {
            const targetTab = btn.getAttribute("data-tab");
            navButtons.forEach(b => b.classList.remove("active"));
            tabContents.forEach(tc => tc.classList.remove("active"));

            btn.classList.add("active");
            document.getElementById(targetTab).classList.add("active");
        });
    });

    // --- CONSTRUTOR DA PILHA DE SWITCHES ---
    const stackBuilderList = document.getElementById("stack-builder-list");
    const btnAddSwitch = document.getElementById("btn-add-switch");

    function renderStackBuilder(members = ["C9300-48U", "C9300-24UX", "C9300-24P"]) {
        stackBuilderList.innerHTML = "";
        members.forEach((model, index) => {
            addSwitchMemberRow(index + 1, model);
        });
    }

    function addSwitchMemberRow(memberNum, selectedModel = "C9300-48U") {
        const item = document.createElement("div");
        item.className = "stack-member-item";
        
        const label = document.createElement("span");
        label.className = "stack-member-label";
        label.textContent = `Switch #${memberNum}`;

        const select = document.createElement("select");
        select.className = "stack-member-select";
        select.name = "stack_models[]";
        
        const models = [
            { val: "C9300-48U", text: "C9300-48U (48 Portas UPOE)" },
            { val: "C9300-24UX", text: "C9300-24UX (24 Portas mGig/Redundante)" },
            { val: "C9300-24P", text: "C9300-24P (24 Portas PoE+)" }
        ];

        models.forEach(m => {
            const opt = document.createElement("option");
            opt.value = m.val;
            opt.textContent = m.text;
            if (m.val === selectedModel) opt.selected = true;
            select.appendChild(opt);
        });

        const btnRemove = document.createElement("button");
        btnRemove.type = "button";
        btnRemove.className = "btn-icon-danger";
        btnRemove.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`;
        btnRemove.addEventListener("click", () => {
            item.remove();
            updateStackMemberLabels();
        });

        item.appendChild(label);
        item.appendChild(select);
        item.appendChild(btnRemove);
        stackBuilderList.appendChild(item);
    }

    function updateStackMemberLabels() {
        const items = stackBuilderList.querySelectorAll(".stack-member-item");
        items.forEach((item, index) => {
            const label = item.querySelector(".stack-member-label");
            if (label) label.textContent = `Switch #${index + 1}`;
        });
    }

    btnAddSwitch.addEventListener("click", () => {
        const currentCount = stackBuilderList.querySelectorAll(".stack-member-item").length;
        addSwitchMemberRow(currentCount + 1, "C9300-48U");
    });

    renderStackBuilder();

    // --- UPLOAD DROP ZONE ---
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("excel_file");
    const fileInfo = document.getElementById("selected-file-info");

    dropZone.addEventListener("click", () => fileInput.click());

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleFileSelected(e.target.files[0]);
        }
    });

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });

    dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));

    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            fileInput.files = e.dataTransfer.files;
            handleFileSelected(file);
        }
    });

    function handleFileSelected(file) {
        selectedFile = file;
        fileInfo.textContent = `📄 Arquivo selecionado: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    }

    // --- PROCESSAMENTO DA FORM ---
    const migrationForm = document.getElementById("migration-form");
    const btnSubmit = document.getElementById("btn-submit");
    const previewTbody = document.getElementById("preview-tbody");
    const metricsGrid = document.getElementById("metrics-grid");
    const downloadContainer = document.getElementById("download-container");
    const btnDownload = document.getElementById("btn-download");
    const resultStatusText = document.getElementById("result-status-text");

    const siteSelectMain = document.getElementById("site_id");
    const defaultVlanDadosInput = document.getElementById("default_vlan_dados");

    siteSelectMain.addEventListener("change", (e) => {
        // Deixa o campo livre para o usuário digitar manualmente
    });

    migrationForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        if (!selectedFile) {
            alert("Por favor, selecione uma planilha .xlsx!");
            return;
        }

        const formData = new FormData(migrationForm);
        formData.append("excel_file", selectedFile);

        btnSubmit.disabled = true;
        btnSubmit.innerHTML = `<span class="spinner"></span> Processando Planilha...`;
        resultStatusText.textContent = "Processando regras e alocando portas...";

        try {
            const resp = await fetch("/api/process-excel", {
                method: "POST",
                body: formData
            });

            const data = await resp.json();

            if (!resp.ok) {
                throw new Error(data.error || "Erro no processamento da planilha.");
            }

            document.getElementById("m-total-migradas").textContent = data.summary.total_migradas;
            document.getElementById("m-total-aps").textContent = data.summary.total_aps;
            document.getElementById("m-total-cams").textContent = data.summary.total_cameras;
            document.getElementById("m-total-geral").textContent = data.summary.total_geral;
            metricsGrid.style.display = "grid";

            btnDownload.href = `/api/download/${data.filename}`;
            downloadContainer.style.display = "block";
            resultStatusText.textContent = `✅ Processamento concluído com sucesso! (${data.summary.total_migradas} linhas migradas)`;

            previewTbody.innerHTML = "";
            data.summary.alocacoes.forEach(item => {
                const tr = document.createElement("tr");

                let badgeClass = "badge-geral";
                if (item.tipo === "AP") badgeClass = "badge-ap";
                if (item.tipo === "Câmera") badgeClass = "badge-cam";
                if (item.tipo === "Novo AP") badgeClass = "badge-novo-ap";
                if (item.is_unidentified) badgeClass = "badge-unidentified";

                if (item.is_unidentified) {
                    tr.style.backgroundColor = "#fff1f2";
                }

                tr.innerHTML = `
                    <td>${item.linha}</td>
                    <td><strong>${item.descricao || 'Dispositivo'}</strong></td>
                    <td><span class="badge ${badgeClass}">${item.is_unidentified ? 'NÃO IDENTIFICADO' : item.tipo}</span></td>
                    <td class="font-mono">${item.switch_novo}</td>
                    <td class="font-mono"><strong>${item.porta_nova}</strong></td>
                    <td><span class="font-mono ${item.is_unidentified ? 'text-danger' : ''}">${item.ip_pool}</span></td>
                    <td><small class="${item.is_unidentified ? 'text-danger fw-bold' : ''}">${item.observacao}</small></td>
                `;
                previewTbody.appendChild(tr);
            });

        } catch (err) {
            alert(`Erro: ${err.message}`);
            resultStatusText.textContent = `❌ Falha no processamento: ${err.message}`;
        } finally {
            btnSubmit.disabled = false;
            btnSubmit.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"></polygon></svg> Processar e Preencher Colunas R a U`;
        }
    });

    // --- GERAR PLANILHA EXEMPLO ---
    const btnGenerateSample = document.getElementById("btn-generate-sample");
    btnGenerateSample.addEventListener("click", async () => {
        try {
            const resp = await fetch("/api/generate-sample", { method: "POST" });
            const data = await resp.json();
            if (data.download_url) {
                window.location.href = data.download_url;
            }
        } catch (err) {
            alert("Erro ao gerar planilha de exemplo.");
        }
    });

    // --- CADASTRO E GERENCIAMENTO DE SITES (ABA 2) ---
    const sitesNavList = document.getElementById("sites-nav-list");
    const configForm = document.getElementById("config-form");
    const editingSiteTitle = document.getElementById("editing-site-title");
    const editingSiteSubtitle = document.getElementById("editing-site-subtitle");
    const btnDeleteSite = document.getElementById("btn-delete-site");

    const vlanDirectList = document.getElementById("vlan-direct-list");
    const vlanReplaceList = document.getElementById("vlan-replace-list");
    const vendorPoolsList = document.getElementById("vendor-pools-list");

    async function loadSiteConfigs(targetSiteIdToSelect = null) {
        try {
            const resp = await fetch("/api/sites");
            siteConfigsCache = await resp.json();

            if (targetSiteIdToSelect && siteConfigsCache[targetSiteIdToSelect]) {
                activeSiteId = targetSiteIdToSelect;
            } else if (!siteConfigsCache[activeSiteId]) {
                activeSiteId = Object.keys(siteConfigsCache)[0] || "default";
            }

            renderSitesNavList();
            renderMainSiteDropdown();
            populateConfigForm(activeSiteId);

        } catch (err) {
            console.error("Erro ao carregar configurações de sites:", err);
        }
    }

    function renderMainSiteDropdown() {
        siteSelectMain.innerHTML = "";
        Object.keys(siteConfigsCache).forEach(sId => {
            const site = siteConfigsCache[sId];
            const opt = document.createElement("option");
            opt.value = sId;
            opt.textContent = site.site_name;
            if (sId === activeSiteId) opt.selected = true;
            siteSelectMain.appendChild(opt);
        });
    }

    function renderSitesNavList() {
        sitesNavList.innerHTML = "";

        Object.keys(siteConfigsCache).forEach(sId => {
            const site = siteConfigsCache[sId];
            const isSelected = (sId === activeSiteId);

            const card = document.createElement("div");
            card.className = `site-item-card ${isSelected ? 'active' : ''}`;
            
            card.innerHTML = `
                <div class="site-item-info">
                    <h4>${site.site_name}</h4>
                    <span>ID: ${site.site_id}</span>
                </div>
                ${sId === 'default' ? '<span class="site-badge-default">PADRÃO</span>' : ''}
            `;

            card.addEventListener("click", () => {
                activeSiteId = sId;
                renderSitesNavList();
                renderMainSiteDropdown();
                populateConfigForm(sId);
            });

            sitesNavList.appendChild(card);
        });
    }

    function populateConfigForm(siteId) {
        const config = siteConfigsCache[siteId] || siteConfigsCache["default"];
        document.getElementById("cfg_site_id").value = config.site_id;
        document.getElementById("cfg_site_name").value = config.site_name;
        document.getElementById("cfg_ap_pool").value = config.ap_pool || "ap_INFRA_VN";
        document.getElementById("cfg_default_vlan_dados").value = config.default_vlan_dados || "";

        editingSiteTitle.textContent = `Editando: ${config.site_name}`;
        editingSiteSubtitle.textContent = `ID do perfil: ${config.site_id}`;

        if (config.site_id === "default") {
            btnDeleteSite.style.display = "none";
        } else {
            btnDeleteSite.style.display = "inline-block";
        }

        // 1. Vendor Pools
        vendorPoolsList.innerHTML = "";
        const vPools = config.vendor_pools || [];
        vPools.forEach(vp => {
            addVendorPoolRow(vp.keywords.join(", "), vp.pool_name, vp.device_type || vp.service_type || "");
        });

        // 2. VLAN Direct
        vlanDirectList.innerHTML = "";
        const vDirect = config.vlan_direct_pools || {};
        Object.keys(vDirect).forEach(vlan => {
            const item = vDirect[vlan];
            if (typeof item === 'object') {
                addVlanDirectRow(vlan, item.pool_name, item.device_type || "");
            } else {
                addVlanDirectRow(vlan, item, "");
            }
        });

        // 3. VLAN Replace
        vlanReplaceList.innerHTML = "";
        const vReplace = config.vlan_replace_pools || {};
        Object.keys(vReplace).forEach(vlan => {
            const item = vReplace[vlan];
            addVlanReplaceRow(vlan, item.replace_vlan, item.pool_name, item.device_type || "");
        });
    }

    function addVendorPoolRow(keywords = "", pool = "notebooks_desktops", deviceType = "") {
        const div = document.createElement("div");
        div.className = "dynamic-row";
        div.innerHTML = `
            <input type="text" placeholder="Palavras-chave Vendor (ex: seiko, epson)" value="${keywords}" class="cfg-vendor-kw">
            <select class="cfg-vendor-pool select-inline">
                <option value="notebooks_desktops" ${pool === 'notebooks_desktops' ? 'selected' : ''}>notebooks_desktops</option>
                <option value="impressoras_plotters_scanners" ${pool === 'impressoras_plotters_scanners' ? 'selected' : ''}>impressoras_plotters_scanners</option>
                <option value="telefoneIP_gateway_ATA" ${pool === 'telefoneIP_gateway_ATA' ? 'selected' : ''}>telefoneIP_gateway_ATA</option>
                <option value="1250_ATA" ${pool === '1250_ATA' ? 'selected' : ''}>1250_ATA</option>
                <option value="ap_INFRA_VN" ${pool === 'ap_INFRA_VN' ? 'selected' : ''}>ap_INFRA_VN</option>
                <option value="outro" ${!['notebooks_desktops','impressoras_plotters_scanners','telefoneIP_gateway_ATA','1250_ATA','ap_INFRA_VN'].includes(pool) ? 'selected' : ''}>[Personalizado]</option>
            </select>
            <input type="text" placeholder="Nome se personalizado" value="${pool}" class="cfg-vendor-pool-custom" style="width: 140px; display: ${!['notebooks_desktops','impressoras_plotters_scanners','telefoneIP_gateway_ATA','1250_ATA','ap_INFRA_VN'].includes(pool) ? 'block' : 'none'};">
            <input type="text" placeholder="Tipo (ex: camera)" value="${deviceType}" class="cfg-vendor-service" style="width: 140px;">
            <button type="button" class="btn-icon-danger btn-remove-row">✕</button>
        `;

        const select = div.querySelector(".cfg-vendor-pool");
        const customInput = div.querySelector(".cfg-vendor-pool-custom");

        select.addEventListener("change", (e) => {
            if (e.target.value === "outro") {
                customInput.style.display = "block";
            } else {
                customInput.style.display = "none";
                customInput.value = e.target.value;
            }
        });

        div.querySelector(".btn-remove-row").addEventListener("click", () => div.remove());
        vendorPoolsList.appendChild(div);
    }

    function addVlanDirectRow(vlan = "", pool = "", deviceType = "") {
        const div = document.createElement("div");
        div.className = "dynamic-row";
        div.innerHTML = `
            <input type="text" placeholder="VLAN ID (108)" value="${vlan}" class="cfg-vlan-direct-key" style="width: 120px;">
            <input type="text" placeholder="Nome do Pool" value="${pool}" class="cfg-vlan-direct-val">
            <input type="text" placeholder="Tipo (ex: camera)" value="${deviceType}" class="cfg-vlan-direct-type" style="width: 140px;">
            <button type="button" class="btn-icon-danger btn-remove-row">✕</button>
        `;
        div.querySelector(".btn-remove-row").addEventListener("click", () => div.remove());
        vlanDirectList.appendChild(div);
    }

    function addVlanReplaceRow(vlanOrig = "", vlanRep = "", pool = "", deviceType = "") {
        const div = document.createElement("div");
        div.className = "dynamic-row";
        div.innerHTML = `
            <input type="text" placeholder="VLAN Origem (748)" value="${vlanOrig}" class="cfg-vlan-rep-orig" style="width: 120px;">
            <input type="text" placeholder="VLAN Nova (1030)" value="${vlanRep}" class="cfg-vlan-rep-new" style="width: 120px;">
            <input type="text" placeholder="Nome do Pool" value="${pool}" class="cfg-vlan-rep-pool">
            <input type="text" placeholder="Tipo (ex: camera)" value="${deviceType}" class="cfg-vlan-rep-type" style="width: 140px;">
            <button type="button" class="btn-icon-danger btn-remove-row">✕</button>
        `;
        div.querySelector(".btn-remove-row").addEventListener("click", () => div.remove());
        vlanReplaceList.appendChild(div);
    }

    document.getElementById("btn-add-vlan-direct").addEventListener("click", () => addVlanDirectRow());
    document.getElementById("btn-add-vlan-replace").addEventListener("click", () => addVlanReplaceRow());
    document.getElementById("btn-add-vendor-pool").addEventListener("click", () => addVendorPoolRow());

    // --- CRIAR NOVO SITE ---
    document.getElementById("btn-new-site").addEventListener("click", async () => {
        const newName = prompt("Informe o Nome do Novo Site (ex: Site Rio de Janeiro):");
        if (!newName || !newName.trim()) return;

        const newId = "site_" + Date.now();
        const newConfig = {
            site_id: newId,
            site_name: newName.trim(),
            ap_pool: "ap_INFRA_VN",
            default_vlan_dados: "",
            vlan_direct_pools: {},
            vlan_replace_pools: {},
            vendor_pools: [
                { keywords: ["dell", "asrock", "lenovo"], pool_name: "notebooks_desktops", device_type: "dispositivo" },
                { keywords: ["seiko", "epson", "argox"], pool_name: "impressoras_plotters_scanners", device_type: "impressora" }
            ]
        };

        try {
            const resp = await fetch(`/api/sites/${newId}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(newConfig)
            });

            if (resp.ok) {
                activeSiteId = newId;
                await loadSiteConfigs(newId);
                alert(`Site "${newName}" criado e salvo com sucesso!`);
            } else {
                alert("Erro ao criar o site.");
            }
        } catch (err) {
            alert("Erro de conexão ao criar site.");
        }
    });

    // --- SALVAR CONFIGURAÇÃO DO SITE ---
    configForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const siteId = document.getElementById("cfg_site_id").value;
        const siteName = document.getElementById("cfg_site_name").value;
        const apPool = document.getElementById("cfg_ap_pool").value;
        const defaultVlanDados = document.getElementById("cfg_default_vlan_dados").value;

        const vlanDirect = {};
        vlanDirectList.querySelectorAll(".dynamic-row").forEach(r => {
            const k = r.querySelector(".cfg-vlan-direct-key").value.trim();
            const v = r.querySelector(".cfg-vlan-direct-val").value.trim();
            const dt = r.querySelector(".cfg-vlan-direct-type") ? r.querySelector(".cfg-vlan-direct-type").value.trim() : "";
            if (k && v) {
                if (dt) {
                    vlanDirect[k] = { pool_name: v, device_type: dt };
                } else {
                    vlanDirect[k] = v;
                }
            }
        });

        const vlanReplace = {};
        vlanReplaceList.querySelectorAll(".dynamic-row").forEach(r => {
            const orig = r.querySelector(".cfg-vlan-rep-orig").value.trim();
            const rep = r.querySelector(".cfg-vlan-rep-new").value.trim();
            const pool = r.querySelector(".cfg-vlan-rep-pool").value.trim();
            const dt = r.querySelector(".cfg-vlan-rep-type") ? r.querySelector(".cfg-vlan-rep-type").value.trim() : "";
            if (orig && rep) {
                vlanReplace[orig] = { 
                    replace_vlan: rep, 
                    pool_name: pool || `POOL_VLAN_${rep}`,
                    device_type: dt
                };
            }
        });

        const vendorPools = [];
        vendorPoolsList.querySelectorAll(".dynamic-row").forEach(r => {
            const kwStr = r.querySelector(".cfg-vendor-kw").value.trim();
            const selPool = r.querySelector(".cfg-vendor-pool").value;
            const customPool = r.querySelector(".cfg-vendor-pool-custom").value.trim();
            const dt = r.querySelector(".cfg-vendor-service").value.trim();

            const poolFinal = (selPool === "outro") ? customPool : selPool;

            if (kwStr && poolFinal) {
                const kws = kwStr.split(",").map(s => s.trim().toLowerCase()).filter(Boolean);
                vendorPools.push({ keywords: kws, pool_name: poolFinal, device_type: dt || "dispositivo" });
            }
        });

        const payload = {
            site_id: siteId,
            site_name: siteName,
            ap_pool: apPool || "ap_INFRA_VN",
            default_vlan_dados: defaultVlanDados || "",
            vlan_direct_pools: vlanDirect,
            vlan_replace_pools: vlanReplace,
            vendor_pools: vendorPools,
            camera_keywords: siteConfigsCache["default"].camera_keywords,
            ap_keywords: siteConfigsCache["default"].ap_keywords
        };

        try {
            const resp = await fetch(`/api/sites/${siteId}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            if (resp.ok) {
                alert(`Configuração do site "${siteName}" salva com sucesso!`);
                await loadSiteConfigs(siteId);
            } else {
                alert("Erro ao salvar configuração.");
            }
        } catch (err) {
            alert("Erro de conexão ao salvar site.");
        }
    });

    // --- EXCLUIR SITE ---
    btnDeleteSite.addEventListener("click", async () => {
        const siteId = document.getElementById("cfg_site_id").value;
        if (siteId === "default") {
            alert("O perfil Padrão não pode ser excluído.");
            return;
        }

        if (confirm(`Tem certeza que deseja excluir o perfil deste site?`)) {
            try {
                const resp = await fetch(`/api/sites/${siteId}`, { method: "DELETE" });
                if (resp.ok) {
                    alert("Site excluído com sucesso!");
                    activeSiteId = "default";
                    await loadSiteConfigs("default");
                }
            } catch (err) {
                alert("Erro ao excluir site.");
            }
        }
    });

    // Carga inicial
    loadSiteConfigs();
});
