# 📱 Carteira Proventos — App Android

Transforma exatamente as telas criadas (Gráfico, Ranking e Cenário Futuro) em um aplicativo Android instalável (.apk).

---

## ✅ Estrutura correta do projeto

Confira se sua pasta no GitHub está EXATAMENTE assim:

```
carteira-proventos/
├── .github/
│   └── workflows/
│       └── build.yml        ← receita de compilação na nuvem
├── src/
│   ├── App.jsx              ← TODO o código das telas
│   ├── main.jsx             ← ponto de entrada
│   └── index.css            ← estilo global
├── index.html
├── package.json
├── vite.config.js
├── capacitor.config.ts
└── .gitignore
```

⚠️ **IMPORTANTE:** `App.jsx`, `main.jsx` e `index.css` precisam estar DENTRO da pasta `src`. Se estiverem soltos na raiz, o build falha.

---

## 🚀 MÉTODO 1 — Compilar na nuvem (recomendado, sem PC)

Este método usa o GitHub Actions. Você não instala nada — a compilação acontece nos servidores do GitHub e você baixa o APK pronto pelo navegador.

### Passo a passo

1. **Suba todos os arquivos** para um repositório no GitHub, respeitando a estrutura acima.

2. Assim que o arquivo `.github/workflows/build.yml` é enviado, a compilação **começa sozinha**.

3. Vá na aba **Actions** do repositório e acompanhe (leva 5-10 minutos).

4. Quando aparecer o **✓ verde**, toque na execução, role até **Artifacts** no rodapé e baixe **carteira-proventos-apk**.

5. Vem um arquivo `.zip` — descompacte e dentro está o `app-debug.apk`.

6. Toque no `.apk` no celular. Aceite "permitir instalação de fontes desconhecidas" e pronto! 🎉

### Se o build falhar (❌ vermelho)

- Toque na execução com erro
- Toque no job **build**
- Procure a etapa com **❌** e toque para expandir
- A mensagem de erro (texto vermelho) diz o que houve

Erros comuns já estão prevenidos neste projeto (Java 21, npm install em vez de npm ci, etc).

---

## 💻 MÉTODO 2 — Compilar no PC com Android Studio

Se preferir gerar localmente:

1. Instale **Node.js 18+** (nodejs.org) e **Android Studio** (developer.android.com/studio)

2. Na pasta do projeto, rode um comando por vez:
```
npm install
npm run build
npx cap add android
npx cap sync android
npx cap open android
```

3. No Android Studio: **Build → Build Bundle(s)/APK(s) → Build APK(s)**

4. Clique em **locate** na notificação para achar o `app-debug.apk`

---

## 🔄 Atualizar o app depois de mudanças

- **Pela nuvem:** edite o arquivo no GitHub e dê commit — o build roda sozinho de novo.
- **No PC:** rode `npm run build` e `npx cap sync android`, depois gere o APK.

---

## 🌐 Testar no navegador antes (opcional)

```
npm install
npm run dev
```
Abre em `http://localhost:5173`.

---

## 📝 Editar os valores de proventos

Os valores de provento por cota e cotação ficam no início de `src/App.jsx`, na constante `ATIVOS`. Edite os campos `prov` (provento por cota) e `cotacao` para atualizar conforme seus dados reais. Após editar, recompile.
