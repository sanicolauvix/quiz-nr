"""
Gera os arquivos PPTX:
  04_NR10.pptx        — Apresentação NR-10
  05_Extintores.pptx  — Apresentação Extintores de Incêndio
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

W = Inches(13.33)
H = Inches(7.5)
BLUE = RGBColor(26, 86, 219)
DARK = RGBColor(17, 24, 39)
MUTED = RGBColor(107, 114, 128)

def new_prs():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    return prs

def add_tb(slide, text, left, top, width, height,
           bold=False, size=20, color=DARK,
           align=PP_ALIGN.LEFT, wrap=True):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = wrap
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return tb

def add_slide(prs, title, body_lines, notes, is_cover=False):
    layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(layout)

    margin = Inches(0.6)
    content_w = W - 2 * margin

    if is_cover:
        # Título centralizado grande
        add_tb(slide, title,
               margin, Inches(2.0), content_w, Inches(1.5),
               bold=True, size=40, color=BLUE, align=PP_ALIGN.CENTER)
        y = Inches(3.7)
        for line in body_lines:
            add_tb(slide, line,
                   margin, y, content_w, Inches(0.55),
                   size=22, color=MUTED, align=PP_ALIGN.CENTER)
            y += Inches(0.55)
    else:
        # Título no topo
        add_tb(slide, title,
               margin, Inches(0.25), content_w, Inches(0.9),
               bold=True, size=26, color=BLUE, align=PP_ALIGN.LEFT)
        # Separador (linha fina via underline não disponível — usamos espaço)
        y = Inches(1.25)
        for line in body_lines:
            add_tb(slide, line,
                   margin, y, content_w, Inches(0.6),
                   size=18, color=DARK)
            y += Inches(0.58)

    # Notas de orador (usadas para narração pelo gerar_videos_MP4.py)
    slide.notes_slide.notes_text_frame.text = notes
    return slide


# ============================================================
# NR-10
# ============================================================
def criar_nr10():
    prs = new_prs()

    slides = [
        dict(
            title="NR-10 — Segurança em Instalações e Serviços em Eletricidade",
            body=[
                "Ministério do Trabalho e Emprego — Portaria n.º 3.214/78",
                "SENAI · Prof. Marcos José dos Santos · Unidade Cariacica",
            ],
            notes=(
                "Bem-vindos à apresentação sobre a NR-10, norma que regulamenta a segurança "
                "em instalações e serviços em eletricidade no Brasil. "
                "Vamos estudar os principais conceitos, perfis profissionais, "
                "sequência de controle e os equipamentos de proteção exigidos."
            ),
            is_cover=True,
        ),
        dict(
            title="O que é a NR-10 e seu Campo de Aplicação",
            body=[
                "• Regulamenta condições mínimas de segurança e saúde para quem trabalha com instalações elétricas.",
                "• Aplica-se a todas as fases: projeto, execução, operação, manutenção, reforma e desmontagem.",
                "• Obrigatória para todo trabalhador exposto a riscos elétricos — não apenas eletricistas.",
                "• Complementada pela NBR 5410 (baixa tensão) e demais normas ABNT aplicáveis.",
                "• Descumprimento: multas, embargos e interdições pelo Ministério do Trabalho.",
            ],
            notes=(
                "A NR-10 se aplica a todas as empresas e trabalhadores que lidem, direta ou indiretamente, "
                "com instalações elétricas. Não é uma norma exclusiva para eletricistas — qualquer pessoa "
                "que possa se expor a riscos elétricos precisa conhecê-la. "
                "Sua base técnica é a NBR 5410, que rege as instalações de baixa tensão."
            ),
        ),
        dict(
            title="Perfis Profissionais (NR-10, item 10.8)",
            body=[
                "• CAPACITADO: recebeu treinamentos de profissional habilitado/qualificado.",
                "           Autorizado para tarefas específicas. Não exige registro em conselho.",
                "• QUALIFICADO: concluiu curso técnico na área elétrica com certificado profissionalizante.",
                "           Não exige registro em conselho de classe.",
                "• HABILITADO: além do curso técnico ou graduação, possui registro no CREA ou CRT.",
                "           Máxima autonomia técnica — pode assinar laudos e projetos.",
                "• PRONTUÁRIO: documento que reúne projetos, laudos e registros da instalação. Obrigação do empregador.",
            ],
            notes=(
                "A NR-10 define três perfis profissionais. O trabalhador capacitado é treinado para tarefas "
                "específicas sem precisar de certificado formal. O qualificado tem certificado profissionalizante. "
                "O habilitado vai além: tem registro em conselho de classe como o CREA. "
                "O prontuário de instalações é documento obrigatório que comprova conformidade com a norma."
            ),
        ),
        dict(
            title="Sequência de Controle — 5 Passos Obrigatórios",
            body=[
                "1. DESENERGIZAÇÃO — desligar todas as fontes (chaves, disjuntores, religadores).",
                "2. SECCIONAMENTO — isolar fisicamente o trecho onde será feito o trabalho.",
                "3. VERIFICAR AUSÊNCIA DE TENSÃO — usar instrumento adequado. NUNCA presuma.",
                "4. ATERRAMENTO PROVISÓRIO — aterrar e curto-circuitar os condutores.",
                "5. SINALIZAÇÃO / BLOQUEIO (LOTO) — cadeado + placa 'Não ligue — homem trabalhando'.",
            ],
            notes=(
                "Antes de qualquer intervenção em instalação elétrica, é obrigatório seguir os cinco passos "
                "da sequência de controle. Primeiro desenergize, depois seccione o trecho, "
                "verifique com instrumento que não há tensão — nunca presuma que está desligado — "
                "faça o aterramento provisório e por fim sinalize com o sistema LOTO: bloqueio com cadeado "
                "e placa de aviso."
            ),
        ),
        dict(
            title="Zonas de Trabalho em Instalações Elétricas",
            body=[
                "🔴 ZONA DE RISCO (ZR): próxima de partes energizadas. Só habilitados com EPIs.",
                "🟡 ZONA CONTROLADA (ZC): ao redor da ZR. Acesso restrito e controlado.",
                "🔵 ZONA DE VIZINHANÇA (ZV): além da ZC. Pessoas informadas dos riscos.",
                "🟢 TRABALHO DESENERGIZADO: condição preferencial — seguir os 5 passos.",
                "⚠️  As distâncias de cada zona variam conforme o nível de tensão (BT, MT ou AT).",
            ],
            notes=(
                "A NR-10 define zonas ao redor de partes energizadas. A Zona de Risco é a mais perigosa, "
                "onde só profissionais habilitados com EPIs adequados podem atuar. "
                "A Zona Controlada exige profissional capacitado ou superior. "
                "A Zona de Vizinhança permite presença de pessoas informadas dos riscos. "
                "Sempre que possível, o trabalho deve ser feito com a instalação desenergizada."
            ),
        ),
        dict(
            title="Trabalho em Instalações Energizadas",
            body=[
                "• É EXCEÇÃO, não regra — só permitido quando a desenergização for inviável.",
                "• Exige Análise de Risco (AR) e emissão de Permissão de Trabalho (PT).",
                "• Realizado por profissional HABILITADO com supervisão e ao menos mais um capacitado.",
                "• EPIs obrigatórios: luvas dielétricas, capacete classe B, óculos, uniforme anti-chama.",
                "• Ferramentas devem ser isoladas/dielétricas, inspecionadas e dentro da validade.",
            ],
            notes=(
                "Trabalhar em instalações energizadas é exceção e exige autorização formal. "
                "Antes de iniciar, é preciso elaborar a Análise de Risco e emitir a Permissão de Trabalho. "
                "O serviço deve ser executado por profissional habilitado, com mais um capacitado presente. "
                "Os equipamentos de proteção dielétricos são obrigatórios e as ferramentas devem estar "
                "isoladas e dentro do prazo de validade do teste."
            ),
        ),
        dict(
            title="EPIs Obrigatórios para Trabalho Elétrico (NR-06 + NR-10)",
            body=[
                "🧤 Luvas dielétricas — classificadas por classe de tensão (0 a 4).",
                "   Inspecionar antes de cada uso: soprar e verificar furos.",
                "⛑️  Capacete classe B (isolante até 20.000 V) com jugular.",
                "   Óculos de proteção contra arco elétrico.",
                "👟 Calçado com solado dielétrico — sem partes metálicas expostas.",
                "👕 Uniforme anti-chama (FR) — proibido tecido sintético (derrete e gruda).",
            ],
            notes=(
                "Os EPIs para trabalho elétrico são determinados pela NR-06 em conjunto com a NR-10. "
                "As luvas dielétricas devem ser classificadas pela classe de tensão e inspecionadas antes "
                "de cada uso. O capacete classe B protege até vinte mil volts. "
                "O calçado deve ter solado dielétrico e o uniforme deve ser anti-chama, "
                "pois tecidos sintéticos derretem e grudam na pele em caso de arco elétrico."
            ),
        ),
        dict(
            title="NBR 5410 e sua Relação com a NR-10",
            body=[
                "NR-10 (trabalhista) — segurança de quem trabalha com eletricidade (pessoas).",
                "NBR 5410 (ABNT)    — projeto e execução de instalações de baixa tensão.",
                "NBR 14039 (ABNT)   — instalações de média tensão.",
                "NBR 5419 (ABNT)    — SPDA, proteção contra descargas atmosféricas.",
                "NBR IEC 60900      — ferramentas isoladas para trabalho energizado.",
                "Conclusão: NR-10 protege o trabalhador; NBR 5410 protege a instalação.",
            ],
            notes=(
                "É importante entender a diferença entre a NR-10 e as normas técnicas da ABNT. "
                "A NR-10 é uma norma trabalhista que protege a integridade do trabalhador. "
                "A NBR 5410 é uma norma técnica que define como projetar e executar instalações elétricas "
                "de baixa tensão. Ambas se complementam: a norma técnica garante a instalação segura "
                "e a NR-10 garante que quem trabalha nela esteja protegido."
            ),
        ),
        dict(
            title="Responsabilidades e Penalidades",
            body=[
                "EMPREGADOR deve: garantir instalações conformes, manter Prontuário atualizado,",
                "  fornecer EPIs gratuitamente, promover treinamentos a cada 2 anos,",
                "  emitir AR e PT, garantir profissional habilitado na supervisão.",
                "TRABALHADOR deve: seguir os 5 passos, usar EPIs corretamente,",
                "  NUNCA presumir desenergização, comunicar condições inseguras,",
                "  exercer o Direito de Recusa em caso de risco grave e iminente.",
                "⚠️  Penalidades: multas (NR-28), embargo, interdição, ações trabalhistas.",
            ],
            notes=(
                "A NR-10 estabelece responsabilidades claras para empregadores e trabalhadores. "
                "O empregador deve fornecer ambiente seguro, EPIs gratuitos e treinamentos periódicos "
                "a cada dois anos. O trabalhador deve seguir os procedimentos, usar os EPIs e "
                "pode exercer o direito de recusa quando houver risco grave e iminente sem sofrer penalidades. "
                "O descumprimento por parte do empregador sujeita à multa, embargo e interdição."
            ),
        ),
        dict(
            title="Resumo — Pontos Essenciais da NR-10",
            body=[
                "✅ Aplica-se a TODOS expostos a riscos elétricos, não só eletricistas.",
                "✅ 3 perfis: Capacitado (treinado) → Qualificado (certificado) → Habilitado (CREA).",
                "✅ 5 passos: Desenergizar → Seccionar → Verificar tensão → Aterrar → Sinalizar.",
                "✅ NUNCA presuma que está desenergizado — teste sempre com instrumento.",
                "✅ Trabalho energizado é EXCEÇÃO: exige AR, PT, habilitado e EPIs dielétricos.",
                "✅ Reciclagem obrigatória a CADA 2 ANOS ou quando houver mudança de risco.",
            ],
            notes=(
                "Vamos revisar os pontos essenciais da NR-10. A norma se aplica a todos os trabalhadores "
                "expostos a riscos elétricos. Existem três perfis profissionais: capacitado, qualificado e habilitado. "
                "Antes de qualquer intervenção, os cinco passos da sequência de controle são obrigatórios. "
                "Nunca presuma que está desenergizado — teste sempre. "
                "Trabalho em instalação energizada é exceção e exige documentação e EPIs específicos. "
                "A reciclagem do treinamento é obrigatória a cada dois anos."
            ),
        ),
    ]

    for s in slides:
        add_slide(prs, s["title"], s["body"], s["notes"],
                  is_cover=s.get("is_cover", False))

    path = "04_NR10.pptx"
    prs.save(path)
    print(f"OK: {path} salvo ({len(slides)} slides)")


# ============================================================
# EXTINTORES
# ============================================================
def criar_extintores():
    prs = new_prs()

    slides = [
        dict(
            title="Extintores de Incêndio — Prevenção e Combate",
            body=[
                "NR-23 · ABNT NBR 12693 · SENAI Unidade Cariacica",
                "Prof. Marcos José dos Santos",
            ],
            notes=(
                "Bem-vindos à aula sobre extintores de incêndio. "
                "Vamos aprender sobre os tipos de incêndio, as classes de fogo, "
                "os agentes extintores corretos para cada situação e como utilizar "
                "um extintor de forma segura e eficiente."
            ),
            is_cover=True,
        ),
        dict(
            title="O Triângulo do Fogo — Como o Incêndio se Forma",
            body=[
                "O fogo precisa de TRÊS elementos simultâneos:",
                "  🔥 COMBUSTÍVEL — material que queima (madeira, papel, gasolina, gás).",
                "  💨 COMBURENTE — oxigênio do ar (mínimo ~16% para sustentar chama).",
                "  ⚡ CALOR (energia de ativação) — fonte de ignição (faísca, chama, atrito).",
                "Remover QUALQUER um dos três elementos extingue o incêndio.",
                "⚠️  Incêndio nunca começa sozinho — identifique e elimine o risco antes.",
            ],
            notes=(
                "O fogo é resultado da combinação de três elementos: combustível, comburente e calor, "
                "formando o chamado triângulo do fogo. Para apagar um incêndio, basta eliminar "
                "um desses elementos. O extintor age principalmente resfriando ou abafando o fogo "
                "e, no caso dos pós químicos, também interrompendo a reação em cadeia."
            ),
        ),
        dict(
            title="Classes de Incêndio",
            body=[
                "CLASSE A — Sólidos com brasas: madeira, papel, tecido, borracha.",
                "           Agente: ÁGUA ou pó ABC. Extinção por resfriamento.",
                "CLASSE B — Líquidos e gases inflamáveis: gasolina, álcool, GLP.",
                "           Agente: PÓ BC ou ABC, CO₂, espuma. NUNCA use água.",
                "CLASSE C — Equipamentos elétricos ENERGIZADOS.",
                "           Agente: PÓ BC ou ABC, CO₂ (não condutores). NUNCA use água.",
                "CLASSE D — Metais combustíveis: magnésio, titânio, sódio.",
                "           Agente: pó especial específico para o metal.",
            ],
            notes=(
                "Os incêndios são classificados em quatro classes principais. "
                "Classe A são os sólidos que formam brasa, como madeira e papel — usa-se água ou pó ABC. "
                "Classe B são líquidos e gases inflamáveis — nunca use água, pois pode espirrar o líquido. "
                "Use pó ou CO2. Classe C são equipamentos elétricos energizados — "
                "nunca use água por risco de choque; use pó seco ou CO2. "
                "Classe D são metais combustíveis que exigem pó especial específico para cada metal."
            ),
        ),
        dict(
            title="Tipos de Extintores e Agentes Extintores",
            body=[
                "💧 ÁGUA PRESSURIZADA — Classe A. Age por resfriamento. Não usar em B, C ou D.",
                "🟡 PÓ QUÍMICO SECO BC — Classes B e C. Interrompe reação em cadeia.",
                "🟠 PÓ QUÍMICO SECO ABC — Classes A, B e C. Versátil, uso mais comum.",
                "❄️  GÁS CARBÔNICO (CO₂) — Classes B e C. Não deixa resíduo. Ideal p/ TI.",
                "🟢 ESPUMA MECÂNICA — Classes A e B. Age por abafamento e resfriamento.",
                "⚠️  Pó ABC deixa resíduo — cause dano ao equipamento eletrônico. Prefira CO₂.",
            ],
            notes=(
                "Existem vários tipos de agentes extintores. O extintor de água age por resfriamento e "
                "serve apenas para a classe A. O pó químico seco BC combate as classes B e C. "
                "O pó ABC é o mais versátil, cobrindo as três classes principais. "
                "O CO2 é ideal para equipamentos eletrônicos pois não deixa resíduo, "
                "mas atenção: em espaços confinados pode causar asfixia. "
                "A espuma é usada para líquidos inflamáveis em grandes volumes."
            ),
        ),
        dict(
            title="Seleção Correta do Extintor por Situação",
            body=[
                "LIXEIRA pegando fogo (papel/madeira)  → Água ou Pó ABC.",
                "Vazamento de gás inflamável (GLP)     → Pó BC ou ABC, CO₂.",
                "Curto em quadro elétrico energizado   → Pó BC, ABC ou CO₂. NUNCA água.",
                "Incêndio em servidor / computadores   → CO₂ (sem resíduo).",
                "Incêndio em depósito de combustível   → Espuma mecânica ou Pó BC.",
                "Regra de ouro: leia sempre o rótulo do extintor antes de usar.",
            ],
            notes=(
                "A escolha do extintor correto é fundamental para a eficácia e segurança. "
                "Para sólidos como papel e madeira, use água ou pó ABC. "
                "Para líquidos e gases inflamáveis, use pó BC, ABC ou CO2, nunca água. "
                "Para equipamentos elétricos energizados, use pó seco ou CO2, nunca água. "
                "Para servidores e computadores, prefira CO2 por não deixar resíduo. "
                "Sempre leia o rótulo do extintor — ele indica as classes para as quais é adequado."
            ),
        ),
        dict(
            title="Como Usar um Extintor — Técnica PASS",
            body=[
                "P — PUXE o pino de segurança (lacre).",
                "A — APONTE o bico/mangueira para a BASE das chamas, não para o topo.",
                "S — APERTE (SQUEEZE) o gatilho com firmeza e de forma contínua.",
                "S — VARRA (SWEEP) da esquerda para a direita, cobrindo toda a base.",
                "Distância segura: entre 1,5 m e 3 m da chama (veja no rótulo).",
                "⚠️  Posicione-se com o vento NAS COSTAS. Tenha sempre rota de fuga.",
            ],
            notes=(
                "A técnica PASS é o método padrão para uso de extintores. "
                "P de puxar o pino de segurança. A de apontar o bico para a base das chamas, "
                "nunca para o topo, pois o fogo começa embaixo. S de apertar o gatilho. "
                "E o segundo S de varrer da esquerda para a direita cobrindo toda a base. "
                "Mantenha a distância indicada no rótulo, fique com o vento nas costas "
                "para o agente extintor não voltar em sua direção e sempre tenha uma rota de fuga."
            ),
        ),
        dict(
            title="Inspeção e Manutenção dos Extintores (NBR 12962)",
            body=[
                "INSPEÇÃO MENSAL (visual) — verificar: lacre intacto, pressão no manômetro (zona verde),",
                "  ausência de danos, corrosão, validade da carga e local de instalação acessível.",
                "MANUTENÇÃO ANUAL — realizada por empresa credenciada pelo INMETRO.",
                "  Verificação interna, recarga e emissão de laudo técnico.",
                "TESTE HIDROSTÁTICO — a cada 5 anos para cilindros de aço.",
                "  Após qualquer uso — mesmo parcial — o extintor deve ser recarregado.",
                "⚠️  Extintor descarregado ou com manômetro na zona vermelha é inoperante.",
            ],
            notes=(
                "Os extintores exigem manutenção regular. A inspeção visual mensal deve ser feita "
                "por funcionário treinado verificando o lacre, a pressão no manômetro — que deve estar "
                "na zona verde — e se o extintor está acessível e sem obstruções. "
                "A manutenção anual deve ser realizada por empresa credenciada pelo INMETRO. "
                "O teste hidrostático do cilindro ocorre a cada cinco anos. "
                "Após qualquer uso, mesmo que parcial, o extintor deve ser imediatamente recarregado."
            ),
        ),
        dict(
            title="Localização, Sinalização e Quantidade (NR-23)",
            body=[
                "ALTURA: parte inferior do extintor a no máximo 1,60 m do piso.",
                "DISTÂNCIA MÁXIMA a percorrer:",
                "  • Classe A: até 15 metros.",
                "  • Classes B e C: até 10 metros.",
                "SINALIZAÇÃO: placa fotoluminescente de identificação obrigatória.",
                "ACESSO: nunca obstruir com móveis, equipamentos ou outros objetos.",
                "QUANTIDADE: calculada conforme área, ocupação e risco do ambiente (NBR 12693).",
            ],
            notes=(
                "A NR-23 e a NBR 12693 definem onde e como posicionar os extintores. "
                "A parte inferior do extintor deve estar a no máximo um metro e sessenta do piso "
                "para facilitar o acesso rápido. A distância máxima a percorrer é de quinze metros "
                "para incêndios classe A e dez metros para as classes B e C. "
                "A sinalização com placa fotoluminescente é obrigatória e o acesso deve ser sempre livre. "
                "A quantidade de extintores é calculada em função da área e do risco do ambiente."
            ),
        ),
        dict(
            title="NR-23 — Proteção Contra Incêndios",
            body=[
                "A NR-23 estabelece obrigações do empregador em proteção contra incêndio:",
                "• Saídas de emergência suficientes e desobstruídas.",
                "• Equipamentos de combate a incêndio em número e tipo adequados.",
                "• Treinamento de todos os trabalhadores no uso dos extintores.",
                "• Brigada de incêndio conforme o porte e risco da empresa.",
                "• Plano de emergência e simulados periódicos.",
                "⚠️  A fiscalização é feita pelo Corpo de Bombeiros e pelo MTE.",
            ],
            notes=(
                "A Norma Regulamentadora número 23 trata da proteção contra incêndios. "
                "O empregador é responsável por manter saídas de emergência desobstruídas, "
                "disponibilizar extintores adequados, treinar todos os funcionários no uso correto "
                "dos equipamentos e manter uma brigada de incêndio conforme o porte da empresa. "
                "Simulados periódicos são obrigatórios para garantir que todos saibam agir "
                "em caso de emergência. A fiscalização é realizada pelo Corpo de Bombeiros e pelo "
                "Ministério do Trabalho e Emprego."
            ),
        ),
        dict(
            title="Resumo — O que Você Precisa Saber sobre Extintores",
            body=[
                "✅ Fogo = Combustível + Comburente + Calor — remova um para extinguir.",
                "✅ 4 classes: A (sólidos) · B (líquidos/gases) · C (elétrico) · D (metais).",
                "✅ NUNCA use água em incêndios classe B ou C.",
                "✅ Técnica PASS: Puxar · Apontar (base) · Apertar · Varrer.",
                "✅ Inspeção mensal + manutenção anual por empresa credenciada.",
                "✅ Distância: 15 m (Cl. A) / 10 m (Cl. B e C) — acesso sempre livre.",
                "✅ Após qualquer uso — recarregue imediatamente.",
            ],
            notes=(
                "Vamos revisar os pontos essenciais sobre extintores de incêndio. "
                "O fogo precisa de três elementos e remover qualquer um deles o extingue. "
                "Lembre as quatro classes de incêndio e nunca use água em fogo elétrico ou de líquidos. "
                "Use a técnica PASS: puxar o pino, apontar para a base, apertar o gatilho e varrer. "
                "Faça inspeção visual mensal e manutenção anual com empresa credenciada. "
                "Mantenha os extintores nas distâncias corretas e o acesso sempre livre. "
                "Após qualquer uso, recarregue imediatamente. Segurança começa na prevenção."
            ),
        ),
    ]

    for s in slides:
        add_slide(prs, s["title"], s["body"], s["notes"],
                  is_cover=s.get("is_cover", False))

    path = "05_Extintores.pptx"
    prs.save(path)
    print(f"OK: {path} salvo ({len(slides)} slides)")


if __name__ == "__main__":
    print("Gerando apresentações PPTX...\n")
    criar_nr10()
    criar_extintores()
    print("\nProntos! Agora execute: python gerar_videos_MP4.py")

