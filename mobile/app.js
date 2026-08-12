const { createApp, ref, onMounted } = Vue;

createApp({
    setup() {
        const isOnline = ref(navigator.onLine);
        const unsyncedDrafts = ref([]);
        const toast = ref({ show: false, message: '' });
        
        // --- CONFIGURAÇÃO DA NUVEM ---
        const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyFjagDBxBuGRxAxvmxrYL1zfgPzXGO7PV8P07ufQqE3fg6xsnNed6sjsgwQ53-1JpUIA/exec"; 
        // -----------------------------
        
        const isAuthenticated = ref(true); // Login desabilitado temporariamente
        const loginForm = ref({ username: '', password: '' });
        let authToken = "";

        const getDefaultVitima = () => ({
            nome: "", documento: "", sexo: "", data_nascimento: "", filicao: "", naturalidade: "", 
            vestes: "", pertences: "", localizacao: "", posicao: "", cabeca: "", membros: "", fenomenos: "",
            lesoes: [""]
        });

        const form = ref({
            num_laudo: "",
            ocorrencia: "",
            autoridade_sel: "",
            data_pericia_input: new Date().toISOString().split('T')[0],
            endereco: "",
            latitude: "",
            longitude: "",
            isolamento: "",
            vitimas: [getDefaultVitima()]
        });

        const showToast = (msg) => {
            toast.value.message = msg;
            toast.value.show = true;
            setTimeout(() => { toast.value.show = false; }, 3000);
        };

        const login = async () => {
            if (loginForm.value.username === 'perito' && loginForm.value.password === 'icrim123') {
                isAuthenticated.value = true;
                authToken = btoa(loginForm.value.username + ":" + loginForm.value.password);
                await localforage.setItem('auth_token', authToken);
                showToast('Login efetuado com sucesso!');
                loadDrafts();
            } else {
                showToast('❌ Usuário ou senha incorretos.');
            }
        };

        const updateNetworkStatus = () => {
            isOnline.value = navigator.onLine;
            if (isOnline.value && isAuthenticated.value) {
                showToast('Conexão restabelecida!');
            }
        };

        const loadDrafts = async () => {
            if (!isAuthenticated.value) return;
            const drafts = await localforage.getItem('laudo_drafts') || [];
            unsyncedDrafts.value = drafts;
        };

        const saveDraft = async () => {
            let drafts = await localforage.getItem('laudo_drafts') || [];
            
            const payload = {
                id: form.value.id || Date.now().toString(), // Keep existing mobile_id or create new
                dados: JSON.parse(JSON.stringify(form.value))
            };
            
            // se já existia, remove o antigo para atualizar
            if (form.value.id) {
                drafts = drafts.filter(d => d.id !== form.value.id);
            }
            
            drafts.unshift(payload); // Add to top
            await localforage.setItem('laudo_drafts', drafts);
            await loadDrafts();
            showToast(form.value.id ? 'Rascunho atualizado! 💾' : 'Rascunho salvo offline! 💾');

            
            // Clear form to start new
            form.value.num_laudo = "";
            form.value.ocorrencia = "";
            form.value.endereco = "";
            form.value.vitimas = [getDefaultVitima()];
            delete form.value.id; // clear draft id
        };

        const loadIntoForm = (draft) => {
            form.value = JSON.parse(JSON.stringify(draft.dados));
            form.value.id = draft.id; // Keep track of the loaded draft
            showToast('Rascunho carregado para edição.');
            window.scrollTo({ top: 0, behavior: 'smooth' });
        };

        const syncData = async () => {
            if (!isOnline.value || unsyncedDrafts.value.length === 0 || !isAuthenticated.value) return;
            
            if (GOOGLE_SCRIPT_URL === "SUA_URL_DO_GOOGLE_SCRIPT_AQUI") {
                showToast("Erro: Configure a URL do Google Script no app.js");
                return;
            }

            showToast('Sincronizando...');
            const drafts = [...unsyncedDrafts.value];
            const failed = [];
            
            for (const draft of drafts) {
                try {
                    // Google Apps Script redirect mode needs follow and no-cors to bypass browser blocks
                    await fetch(GOOGLE_SCRIPT_URL, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: {
                            'Content-Type': 'text/plain;charset=utf-8'
                        },
                        body: JSON.stringify({
                            key: "perito:icrim123", // A senha configurada
                            mobile_id: draft.id,
                            dados: draft.dados
                        })
                    });
                    // Em mode 'no-cors', a resposta é opaca, então assumimos sucesso se a rede não falhou.
                } catch(e) {
                    console.error("Falha ao sincronizar: ", e);
                    failed.push(draft);
                }
            }
            
            await localforage.setItem('laudo_drafts', failed);
            await loadDrafts();
            
            if (failed.length === 0) {
                showToast('Todos sincronizados com sucesso! 🚀');
            } else {
                showToast(`Falha em ${failed.length} laudos. Tente depois.`);
            }
        };

        const getGPS = () => {
            if (navigator.geolocation) {
                showToast('Buscando localização...');
                navigator.geolocation.getCurrentPosition(
                    pos => {
                        form.value.latitude = pos.coords.latitude.toFixed(6);
                        form.value.longitude = pos.coords.longitude.toFixed(6);
                        showToast('Localização obtida! 📍');
                    },
                    err => {
                        showToast('Erro ao acessar GPS.');
                    },
                    { enableHighAccuracy: true }
                );
            }
        };

        const adicionarVitima = () => { form.value.vitimas.push(getDefaultVitima()); };
        const removerVitima = (idx) => { form.value.vitimas.splice(idx, 1); };
        
        const adicionarLesao = (vIdx) => { form.value.vitimas[vIdx].lesoes.push(""); };
        const removerLesao = (vIdx, lIdx) => { form.value.vitimas[vIdx].lesoes.splice(lIdx, 1); };

        onMounted(async () => {
            window.addEventListener('online', updateNetworkStatus);
            window.addEventListener('offline', updateNetworkStatus);
            
            const token = await localforage.getItem('auth_token');
            if (token) {
                authToken = token;
                isAuthenticated.value = true;
                loadDrafts();
            }
        });

        return {
            form, isOnline, unsyncedDrafts, toast,
            isAuthenticated, loginForm, login,
            saveDraft, syncData, getGPS, loadIntoForm,
            adicionarVitima, removerVitima,
            adicionarLesao, removerLesao
        }
    }
}).mount('#app');
