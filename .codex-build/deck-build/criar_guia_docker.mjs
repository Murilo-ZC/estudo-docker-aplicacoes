import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const SKILL_DIR = "/Users/murilo/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations";
const workspaceDir = "/Users/murilo/Downloads/estudo-docker-aplicacoes";
const TMP_DIR = path.join(workspaceDir, ".codex-build", "deck-guide");
const FINAL_PPTX = path.join(workspaceDir, "materiais", "guia_estudos_docker_imagens_compose_final.pptx");
const RUNTIME_PYTHON = "/Users/murilo/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";
const referencePath = "/Users/murilo/Downloads/Encontro 01 - Containers e Docker.pptx";

const { makeNativeBulletParagraphs, finalizePresentation } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools", "artifact_tool_utils.mjs")).href,
);

await fs.mkdir(TMP_DIR, { recursive: true });
await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });

const W = 2560;
const H = 1440;
const C = {
  white: "#FFFFFF",
  ink: "#2D2D32",
  muted: "#68686F",
  purple: "#4F3292",
  purpleDark: "#35234A",
  palePurple: "#F2EEF9",
  paleGreen: "#EEF6F1",
  paleGray: "#F3F3F5",
  code: "#24242E",
  codeText: "#F2F2F4",
  gray: "#686868",
  red: "#C53A3E",
};
const BODY = "Montserrat";
const MONO = "Space Mono";
const CODE = "Space Mono";

const presentation = Presentation.create({ slideSize: { width: W, height: H } });

function addBox(slide, name, x, y, w, h, fill = "none", line = "none", radius = false) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: { style: "solid", fill: line, width: line === "none" ? 0 : 1 },
  });
}

function addText(slide, name, text, x, y, w, h, opts = {}) {
  const shape = addBox(slide, name, x, y, w, h, opts.fill ?? "none", opts.line ?? "none", opts.radius ?? false);
  shape.text = text;
  shape.text.style = {
    typeface: opts.typeface ?? BODY,
    fontSize: opts.fontSize ?? 30,
    bold: opts.bold ?? false,
    color: opts.color ?? C.ink,
    alignment: opts.alignment ?? "left",
    verticalAlignment: opts.verticalAlignment ?? "top",
    autoFit: opts.autoFit ?? "shrinkText",
  };
  return shape;
}

function addTitle(slide, title, number) {
  addText(slide, `title-${number}`, title, 128, 70, 2304, 120, {
    typeface: MONO, fontSize: 48, bold: true, color: C.purple,
  });
  addText(slide, `number-${number}`, String(number), 1220, 1360, 120, 42, {
    fontSize: 20, color: C.muted, alignment: "center", verticalAlignment: "middle",
  });
}

function addBullets(slide, name, items, x, y, w, h, opts = {}) {
  const shape = addBox(slide, name, x, y, w, h);
  shape.text = makeNativeBulletParagraphs(items, {
    marginLeftPoints: opts.marginLeftPoints ?? 20,
    hangingPoints: opts.hangingPoints ?? 10,
    spaceAfterPoints: opts.spaceAfterPoints ?? 12,
  });
  shape.text.style = {
    typeface: opts.typeface ?? BODY,
    fontSize: opts.fontSize ?? 30,
    color: opts.color ?? C.ink,
    autoFit: "shrinkText",
  };
  return shape;
}

function addCallout(slide, text, color = C.palePurple) {
  addText(slide, `callout-${slide.id}`, text, 128, 1190, 2304, 115, {
    fill: color, radius: true, fontSize: 25, bold: true, color: C.purple,
    verticalAlignment: "middle",
  });
}

function addCode(slide, name, code, x, y, w, h, fontSize = 24) {
  return addText(slide, name, code, x, y, w, h, {
    fill: C.code, radius: true, typeface: CODE, fontSize, color: C.codeText,
    verticalAlignment: "middle",
  });
}

function addSection(title, subtitle, number) {
  const slide = presentation.slides.add();
  slide.background.fill = C.gray;
  addText(slide, `section-title-${number}`, title, 230, 500, 2100, 150, {
    fontSize: 62, bold: true, color: C.white, alignment: "center", verticalAlignment: "middle",
  });
  addText(slide, `section-subtitle-${number}`, subtitle, 350, 690, 1860, 80, {
    fontSize: 28, color: C.white, alignment: "center", verticalAlignment: "middle",
  });
  addText(slide, `section-number-${number}`, String(number), 1220, 1360, 120, 42, {
    fontSize: 20, color: C.white, alignment: "center", verticalAlignment: "middle",
  });
  return slide;
}

function note(slide, text) {
  slide.speakerNotes.textFrame.setText(text);
}

// 1 Capa
{
  const slide = presentation.slides.add();
  slide.background.fill = C.gray;
  addText(slide, "cover-title", "Guia de estudos\nImagens e Docker Compose", 210, 520, 1600, 230, {
    fontSize: 64, bold: true, color: C.white, verticalAlignment: "middle",
  });
  addText(slide, "cover-subtitle", "Conceitos essenciais e comandos de uso", 214, 775, 1450, 70, {
    fontSize: 30, color: C.white,
  });
  addText(slide, "cover-course", "Containers e Docker", 214, 1070, 900, 55, {
    fontSize: 24, color: C.white,
  });
  addText(slide, "cover-number", "1", 1220, 1360, 120, 42, {
    fontSize: 20, color: C.white, alignment: "center", verticalAlignment: "middle",
  });
}

// 2
{
  const slide = presentation.slides.add(); addTitle(slide, "Objetivos deste guia", 2);
  addBullets(slide, "objectives", [
    "Distinguir imagem de container",
    "Ler e escrever um Dockerfile pequeno",
    "Construir uma imagem e iniciar containers",
    "Descrever dois serviços com Docker Compose",
    "Usar portas, nomes de serviço e logs para testar a solução",
  ], 170, 290, 2050, 650, { fontSize: 36, spaceAfterPoints: 17 });
  addCallout(slide, "O objetivo é entender as decisões. Os comandos vêm depois do modelo mental.");
  note(slide, "Guia curto para revisão. Os casos do repositório oferecem material de prática após a apresentação.");
}

// 3
{
  const slide = presentation.slides.add(); addTitle(slide, "Imagem container e registry", 3);
  const reg = addText(slide, "registry", "REGISTRY\nrepositório de imagens", 210, 450, 560, 230, {
    fill: C.paleGray, radius: true, fontSize: 31, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle",
  });
  const img = addText(slide, "image", "IMAGEM\nmodelo somente leitura", 1000, 450, 560, 230, {
    fill: C.palePurple, radius: true, fontSize: 31, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle",
  });
  const cont = addText(slide, "container", "CONTAINER\ninstância em execução", 1790, 450, 560, 230, {
    fill: C.paleGreen, radius: true, fontSize: 31, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle",
  });
  slide.shapes.connect(reg, img, { kind: "straight", fromSide: "right", toSide: "left", line: { style: "solid", fill: C.purple, width: 5 }, tail: { type: "arrow", width: "med", length: "med" } });
  slide.shapes.connect(img, cont, { kind: "straight", fromSide: "right", toSide: "left", line: { style: "solid", fill: C.purple, width: 5 }, tail: { type: "arrow", width: "med", length: "med" } });
  addText(slide, "flow-caption", "docker pull baixa uma imagem. docker run cria um container novo a partir dela.", 340, 800, 1880, 90, {
    fontSize: 31, alignment: "center",
  });
  addCallout(slide, "Uma imagem pode gerar vários containers com nomes, portas e configurações diferentes.");
}

// 4
{
  const slide = presentation.slides.add(); addTitle(slide, "A imagem é formada por camadas", 4);
  const layers = [
    ["Imagem base", "python:3.12-slim", C.paleGray, 420],
    ["Dependências", "pip install -r requirements.txt", C.palePurple, 375],
    ["Código da aplicação", "COPY . .", C.paleGreen, 330],
    ["Metadados de execução", "EXPOSE e CMD", "#F8EEEE", 285],
  ];
  let y = 330;
  for (const [label, detail, fill, inset] of layers) {
    addText(slide, `layer-${label}`, `${label}\n${detail}`, 350 + inset / 2, y, 1860 - inset, 155, {
      fill, radius: true, fontSize: 28, bold: true, color: C.ink, alignment: "center", verticalAlignment: "middle",
    });
    y += 175;
  }
  addCallout(slide, "Mudanças próximas ao fim do Dockerfile reaproveitam mais camadas do build anterior.");
}

// 5
{
  const slide = presentation.slides.add(); addTitle(slide, "O Dockerfile descreve a imagem", 5);
  const items = [
    ["FROM", "define a imagem base"],
    ["WORKDIR", "escolhe o diretório de trabalho"],
    ["COPY", "leva arquivos para a imagem"],
    ["RUN", "executa uma etapa durante o build"],
    ["EXPOSE", "documenta a porta usada pelo processo"],
    ["CMD", "define o processo principal do container"],
  ];
  let y = 270;
  for (let i = 0; i < items.length; i++) {
    const [key, desc] = items[i];
    addText(slide, `key-${key}`, key, 170, y, 370, 105, {
      fill: i % 2 === 0 ? C.palePurple : C.paleGray, radius: true, typeface: MONO,
      fontSize: 28, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle",
    });
    addText(slide, `desc-${key}`, desc, 600, y, 1600, 105, { fontSize: 30, verticalAlignment: "middle" });
    y += 135;
  }
}

// 6
{
  const slide = presentation.slides.add(); addTitle(slide, "Exemplo de Dockerfile", 6);
  addCode(slide, "dockerfile-code", `FROM python:3.12-slim\n\nWORKDIR /app\n\nCOPY requirements.txt .\nRUN pip install --no-cache-dir -r requirements.txt\n\nCOPY . .\n\nEXPOSE 5050\nCMD ["gunicorn", "--bind", "0.0.0.0:5050", "main:app"]`, 150, 245, 1440, 860, 25);
  addBullets(slide, "dockerfile-notes", [
    "O build cria a imagem",
    "RUN acontece durante o build",
    "CMD acontece quando o container inicia",
    "EXPOSE não publica a porta no computador",
  ], 1690, 320, 720, 620, { fontSize: 28, spaceAfterPoints: 14 });
  addCallout(slide, "Copiar as dependências antes do código ajuda a aproveitar o cache quando só o código muda.");
  note(slide, "Exemplo alinhado aos Dockerfiles presentes nos casos do repositório.");
}

// 7
{
  const slide = presentation.slides.add(); addTitle(slide, "Construção e identificação da imagem", 7);
  addCode(slide, "build-code", `docker build -t api-recados:1.0 .\n\ndocker image ls\n\ndocker image inspect api-recados:1.0`, 180, 320, 1500, 500, 29);
  addText(slide, "tag-title", "Nome e tag", 1770, 335, 620, 70, { typeface: MONO, fontSize: 30, bold: true, color: C.purple });
  addText(slide, "tag-copy", "api-recados identifica o projeto.\n\n1.0 identifica uma versão.\n\nUma alteração no Dockerfile ou no código exige um novo build.", 1770, 430, 620, 390, { fontSize: 29 });
  addCallout(slide, "A tag latest não explica qual versão está em uso. Tags explícitas facilitam testes e retorno.");
}

// 8
{
  const slide = presentation.slides.add(); addTitle(slide, "Imagem e container têm papéis diferentes", 8);
  addText(slide, "image-side", "IMAGEM", 190, 310, 970, 90, { typeface: MONO, fontSize: 38, bold: true, color: C.purple, alignment: "center" });
  addBullets(slide, "image-list", ["Resultado do build", "Camadas somente leitura", "Pode ser enviada a um registry", "Não representa uma execução"], 240, 460, 870, 500, { fontSize: 31 });
  addText(slide, "container-side", "CONTAINER", 1400, 310, 970, 90, { typeface: MONO, fontSize: 38, bold: true, color: C.purple, alignment: "center" });
  addBullets(slide, "container-list", ["Criado a partir da imagem", "Possui nome e estado", "Executa um processo principal", "Pode publicar portas e receber variáveis"], 1450, 460, 870, 500, { fontSize: 31 });
  addCallout(slide, "Alterar um arquivo dentro do container não atualiza a imagem que o criou.");
}

// 9
{
  const slide = presentation.slides.add(); addTitle(slide, "Criação do container e publicação de porta", 9);
  addCode(slide, "run-code", "docker run -d --name recados -p 9100:5050 api-recados:1.0", 180, 280, 2200, 150, 29);
  const host = addText(slide, "host-port", "COMPUTADOR\nlocalhost:9100", 260, 610, 700, 220, { fill: C.paleGray, radius: true, fontSize: 32, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle" });
  const cont = addText(slide, "container-port", "CONTAINER\nprocesso em 5050", 1600, 610, 700, 220, { fill: C.paleGreen, radius: true, fontSize: 32, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle" });
  slide.shapes.connect(host, cont, { kind: "straight", fromSide: "right", toSide: "left", line: { style: "solid", fill: C.purple, width: 6 }, tail: { type: "arrow", width: "med", length: "med" } });
  addText(slide, "port-map", "-p 9100:5050", 1010, 650, 540, 90, { typeface: MONO, fontSize: 29, bold: true, color: C.purple, alignment: "center" });
  addCallout(slide, "A porta à esquerda pertence ao computador. A porta à direita pertence ao container.");
}

// 10
{
  const slide = presentation.slides.add(); addTitle(slide, "Comandos para observar e controlar", 10);
  addCode(slide, "lifecycle-code", `docker ps\ndocker logs -f recados\ndocker exec -it recados sh\ndocker stop recados\ndocker start recados\ndocker rm -f recados`, 180, 280, 1180, 700, 28);
  addText(slide, "lifecycle-title", "Perguntas para o diagnóstico", 1510, 300, 800, 70, { typeface: MONO, fontSize: 30, bold: true, color: C.purple });
  addBullets(slide, "lifecycle-questions", [
    "O container está em execução?",
    "Qual porta foi publicada?",
    "O processo iniciou sem erro?",
    "A requisição chegou ao serviço?",
  ], 1510, 420, 800, 500, { fontSize: 31 });
  addCallout(slide, "O log mostra o que o processo escreveu. Ele costuma ser a primeira evidência útil.");
}

// 11 divisor
addSection("Docker Compose", "Vários serviços descritos em um único arquivo", 11);

// 12
{
  const slide = presentation.slides.add(); addTitle(slide, "O problema que o Compose resolve", 12);
  addBullets(slide, "compose-problem", [
    "Uma aplicação pode depender de frontend, API e banco de dados",
    "Cada serviço precisa de imagem, configuração, rede e armazenamento",
    "Uma sequência longa de docker run é difícil de repetir e revisar",
    "O arquivo Compose registra como os serviços trabalham juntos",
  ], 180, 300, 2100, 650, { fontSize: 35, spaceAfterPoints: 18 });
  addCallout(slide, "Compose descreve a execução. Dockerfiles descrevem como construir as imagens.");
}

// 13
{
  const slide = presentation.slides.add(); addTitle(slide, "Estrutura de um arquivo Compose", 13);
  addCode(slide, "compose-structure", `services:\n  site:\n    build: ./frontend\n    ports:\n      - "8080:80"\n\n  api:\n    build: ./backend`, 160, 270, 1080, 750, 31);
  addText(slide, "compose-fields-title", "Campos mais usados", 1400, 285, 900, 70, { typeface: MONO, fontSize: 31, bold: true, color: C.purple });
  addBullets(slide, "compose-fields", [
    "services define os containers",
    "build indica onde está o Dockerfile",
    "image usa uma imagem já construída",
    "ports publica uma porta no computador",
    "environment envia configurações",
    "volumes preserva ou compartilha arquivos",
  ], 1400, 400, 920, 620, { fontSize: 29, spaceAfterPoints: 11 });
}

// 14
{
  const slide = presentation.slides.add(); addTitle(slide, "Compose com frontend e backend", 14);
  addCode(slide, "compose-full", `services:\n  frontend:\n    build: ./frontend\n    ports:\n      - "8080:80"\n    depends_on:\n      - backend\n\n  backend:\n    build: ./backend`, 170, 245, 1250, 830, 29);
  addBullets(slide, "compose-full-notes", [
    "O usuário acessa somente o frontend",
    "O backend fica disponível na rede interna",
    "O nome backend funciona como endereço interno",
    "A porta interna vem do processo no container",
  ], 1530, 340, 820, 600, { fontSize: 30, spaceAfterPoints: 15 });
  addCallout(slide, "O mapeamento 8080:80 não muda a porta usada entre os serviços.");
  note(slide, "Estrutura equivalente ao caso03, com foco no papel de cada campo e sem entregar os desafios propostos.");
}

// 15
{
  const slide = presentation.slides.add(); addTitle(slide, "Rede interna e nomes de serviço", 15);
  const browser = addText(slide, "browser", "NAVEGADOR\nlocalhost:8080", 150, 520, 600, 210, { fill: C.paleGray, radius: true, fontSize: 30, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle" });
  const front = addText(slide, "front", "FRONTEND\nNginx na porta 80", 980, 520, 600, 210, { fill: C.palePurple, radius: true, fontSize: 30, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle" });
  const back = addText(slide, "back", "BACKEND\nbackend:5050", 1810, 520, 600, 210, { fill: C.paleGreen, radius: true, fontSize: 30, bold: true, color: C.purple, alignment: "center", verticalAlignment: "middle" });
  slide.shapes.connect(browser, front, { kind: "straight", fromSide: "right", toSide: "left", line: { style: "solid", fill: C.purple, width: 5 }, tail: { type: "arrow", width: "med", length: "med" } });
  slide.shapes.connect(front, back, { kind: "straight", fromSide: "right", toSide: "left", line: { style: "solid", fill: C.purple, width: 5 }, tail: { type: "arrow", width: "med", length: "med" } });
  addText(slide, "public-label", "porta publicada", 735, 770, 470, 60, { fontSize: 24, color: C.muted, alignment: "center" });
  addText(slide, "internal-label", "rede do Compose", 1560, 770, 470, 60, { fontSize: 24, color: C.muted, alignment: "center" });
  addCallout(slide, "O navegador não conhece o nome backend. O Nginx usa esse nome dentro da rede do Compose.");
}

// 16
{
  const slide = presentation.slides.add(); addTitle(slide, "Proxy reverso no frontend", 16);
  addCode(slide, "nginx-code", `location /api/ {\n    proxy_pass http://backend:5050/;\n    proxy_set_header Host $host;\n    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;\n}`, 160, 310, 1450, 500, 27);
  addText(slide, "proxy-example-title", "Exemplo de tradução", 1740, 330, 650, 70, { typeface: MONO, fontSize: 30, bold: true, color: C.purple });
  addText(slide, "proxy-example", "/api/rotas\n\nchega ao backend como\n\n/rotas", 1740, 455, 650, 360, { fontSize: 34, bold: true, alignment: "center", verticalAlignment: "middle" });
  addCallout(slide, "A barra final de proxy_pass muda o caminho encaminhado. A porta deve coincidir com a porta interna do backend.");
}

// 17
{
  const slide = presentation.slides.add(); addTitle(slide, "Dependências volumes e configuração", 17);
  const cols = [
    ["depends_on", "Registra dependência e ordem de início. Sozinho, não garante que a aplicação já esteja pronta."],
    ["volumes", "Mantém dados fora da camada gravável do container ou compartilha arquivos necessários."],
    ["environment", "Entrega valores que variam entre ambientes sem criar outra imagem."],
  ];
  let x = 150;
  for (let i = 0; i < cols.length; i++) {
    const [head, copy] = cols[i];
    addText(slide, `compose-concept-${head}`, head, x, 330, 700, 100, { typeface: MONO, fontSize: 31, bold: true, color: C.purple, alignment: "center" });
    addText(slide, `compose-copy-${head}`, copy, x, 470, 700, 390, { fill: i === 1 ? C.paleGreen : C.palePurple, radius: true, fontSize: 29, alignment: "center", verticalAlignment: "middle" });
    x += 825;
  }
  addCallout(slide, "Quando prontidão importa, combine healthcheck com uma condição de dependência adequada.");
}

// 18
{
  const slide = presentation.slides.add(); addTitle(slide, "Fluxo de uso do Compose", 18);
  addCode(slide, "compose-commands", `docker compose config\ndocker compose up -d --build\ndocker compose ps\ndocker compose logs -f\ndocker compose down`, 180, 300, 1250, 650, 30);
  addText(slide, "workflow-title", "Leitura do fluxo", 1570, 315, 760, 70, { typeface: MONO, fontSize: 31, bold: true, color: C.purple });
  addBullets(slide, "workflow-list", [
    "Validar o arquivo",
    "Construir e iniciar os serviços",
    "Conferir o estado",
    "Acompanhar as mensagens",
    "Encerrar a solução",
  ], 1570, 430, 760, 520, { fontSize: 31, spaceAfterPoints: 14 });
  addCallout(slide, "Use --build quando o Dockerfile ou os arquivos copiados para a imagem mudarem.");
}

// 19
{
  const slide = presentation.slides.add(); addTitle(slide, "Diagnóstico da comunicação entre serviços", 19);
  addText(slide, "diagnosis-question", "Quando o frontend abre e a API não responde", 170, 250, 2200, 80, { fontSize: 34, bold: true, color: C.ink });
  addBullets(slide, "diagnosis-list", [
    "Confirme se todos os serviços aparecem em docker compose ps",
    "Leia os logs do frontend e do backend",
    "Compare a porta em que o processo escuta com a porta usada pelo proxy",
    "Use o nome do serviço para a comunicação interna",
    "Teste primeiro o backend e depois o caminho público do frontend",
  ], 210, 390, 2100, 650, { fontSize: 32, spaceAfterPoints: 16 });
  addCallout(slide, "HTTP 502 costuma indicar que o proxy recebeu a requisição, mas não conseguiu falar com o serviço de destino.", "#F8EEEE");
  note(slide, "O caso03 do repositório permite praticar este roteiro com uma incompatibilidade de porta no proxy.");
}

// 20
{
  const slide = presentation.slides.add(); addTitle(slide, "Checklist de estudo", 20);
  addBullets(slide, "study-checklist", [
    "Consigo explicar a diferença entre imagem e container",
    "Consigo localizar FROM, RUN e CMD em um Dockerfile",
    "Consigo construir uma imagem com nome e tag",
    "Consigo publicar uma porta e interpretar host:container",
    "Consigo descrever dois serviços em um arquivo Compose",
    "Consigo explicar como um serviço encontra outro pela rede interna",
    "Consigo usar ps e logs para investigar uma falha",
  ], 180, 260, 2130, 820, { fontSize: 32, spaceAfterPoints: 14 });
  addCallout(slide, "Próximo passo: resolver os desafios sem copiar os arquivos dos casos.");
}

const requirements = {
  explicitTotalSlideCount: 20,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
};
const fontPolicy = {
  basis: "reference",
  families: [MONO, BODY],
  referencePath,
  referenceSha256: "a4fdca7ed19a298f65308ea07f02fc743c0d2d4d84ca28f55008f15da0336935",
};
const expectedSlideSizeEmu = "24384000,13716000";
const stagingDir = path.join(workspaceDir, ".codex-build", "deck-finalizer");
await fs.mkdir(stagingDir, { recursive: true });
const candidatePath = path.join(stagingDir, "guia-candidate-final.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);

const result = await finalizePresentation({
  ...requirements,
  workspaceDir,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools", "inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools", "inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", expectedSlideSizeEmu, "--validate-bullet-geometry", "--validate-heading-fit"],
  requiredNativeTableOwnerSlides: [],
  fontPolicy,
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, "guia_estudos_docker_final.validation.json"),
});

console.log(JSON.stringify({ finalPath: FINAL_PPTX, result }, null, 2));
