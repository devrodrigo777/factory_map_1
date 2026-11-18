import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Paleta Industrial Knauf / EasyLOTOTO (Moderna, nítida e técnica)
    C_DARK_BG = RGBColor(11, 19, 30)       # #0B131E (Fundo escuro capa)
    C_HEADER_BG = RGBColor(15, 23, 42)     # #0F172A (Header e barras)
    C_CYAN = RGBColor(2, 132, 199)         # #0284C7 (Azul Primário Knauf)
    C_CYAN_LIGHT = RGBColor(56, 189, 248)  # #38BDF8 (Ciano Vivo de Destaque)
    C_ORANGE = RGBColor(234, 88, 12)       # #EA580C (Laranja Alerta LOTOTO)
    C_GREEN = RGBColor(34, 197, 94)        # #22C55E (Verde Conforme)
    C_WHITE = RGBColor(255, 255, 255)
    C_CARD_BG = RGBColor(255, 255, 255)
    C_CARD_BORDER = RGBColor(226, 232, 240)# #E2E8F0 (Borda sutil 1px)
    C_PAGE_BG = RGBColor(248, 250, 252)    # #F8FAFC (Fundo das páginas)
    C_TEXT_TITLE = RGBColor(15, 23, 42)    # #0F172A
    C_TEXT_BODY = RGBColor(51, 65, 85)      # #334155
    C_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B
    C_DICA_BG = RGBColor(240, 249, 255)    # #F0F9FF (Fundo dica azul claro)
    C_DICA_BORDER = RGBColor(186, 230, 253)# #BAE6FD

    def add_bg(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_header(slide, tag_text, title_text, page_num):
        # Header escuro linear no topo
        header_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.15))
        header_bar.fill.solid()
        header_bar.fill.fore_color.rgb = C_HEADER_BG
        header_bar.line.fill.background()

        # Linha técnica de acento ciano/azul Knauf
        accent_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(1.13), Inches(13.333), Inches(0.04))
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = C_CYAN
        accent_line.line.fill.background()

        # Texto da categoria e título
        tb = slide.shapes.add_textbox(Inches(0.75), Inches(0.12), Inches(10.5), Inches(0.95))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_tag = tf.paragraphs[0]
        p_tag.text = tag_text.upper()
        p_tag.font.name = 'Segoe UI'
        p_tag.font.size = Pt(9.5)
        p_tag.font.bold = True
        p_tag.font.color.rgb = C_CYAN_LIGHT

        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.name = 'Segoe UI'
        p_title.font.size = Pt(18)
        p_title.font.bold = True
        p_title.font.color.rgb = C_WHITE

        # Indicador de Página no canto direito
        tb_page = slide.shapes.add_textbox(Inches(11.2), Inches(0.35), Inches(1.4), Inches(0.5))
        tf_page = tb_page.text_frame
        tf_page.margin_right = Inches(0.75)
        p_num = tf_page.paragraphs[0]
        p_num.text = f"{page_num:02d} / 16"
        p_num.font.name = 'Segoe UI'
        p_num.font.size = Pt(13)
        p_num.font.bold = True
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 1: CAPA INSTITUCIONAL TÉCNICA
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    add_bg(slide1, C_DARK_BG)

    # Faixa lateral decorativa azul
    side_stripe = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.35), Inches(7.5))
    side_stripe.fill.solid()
    side_stripe.fill.fore_color.rgb = C_CYAN
    side_stripe.line.fill.background()

    # Badge de Categoria (Retângulo limpo e técnico)
    badge = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.05), Inches(3.6), Inches(0.38))
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor(30, 41, 59)
    badge.line.color.rgb = C_CYAN_LIGHT
    badge.line.width = Pt(1)
    tf_b = badge.text_frame
    p_b = tf_b.paragraphs[0]
    p_b.text = "MANUAL VISUAL DE TREINAMENTO • V2.0"
    p_b.font.name = 'Segoe UI'
    p_b.font.size = Pt(9.5)
    p_b.font.bold = True
    p_b.alignment = PP_ALIGN.CENTER
    p_b.font.color.rgb = C_CYAN_LIGHT

    # Título Principal
    tb_title = slide1.shapes.add_textbox(Inches(1.2), Inches(1.65), Inches(11.0), Inches(2.2))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True

    p_t1 = tf_t.paragraphs[0]
    p_t1.text = "EasyLOTOTO for Knauf v2"
    p_t1.font.name = 'Segoe UI'
    p_t1.font.size = Pt(40)
    p_t1.font.bold = True
    p_t1.font.color.rgb = C_WHITE

    p_t2 = tf_t.add_paragraph()
    p_t2.text = "Guia Passo a Passo de Implantação, Configuração e Operação"
    p_t2.font.name = 'Segoe UI'
    p_t2.font.size = Pt(20)
    p_t2.font.color.rgb = RGBColor(203, 213, 225)
    p_t2.space_before = Pt(8)

    p_t3 = tf_t.add_paragraph()
    p_t3.text = "Gestão visual de bloqueios e etiquetagens de energia perigosa (Lockout, Tagout, Tryout) na fábrica."
    p_t3.font.name = 'Segoe UI'
    p_t3.font.size = Pt(13)
    p_t3.font.color.rgb = RGBColor(148, 163, 184)
    p_t3.space_before = Pt(8)

    # 4 Cards Informativos de Pilares (Retângulos limpos com bordas técnicas)
    pilares = [
        ("Operação Offline", "Servidor autônomo local e SQLite WAL com execução via atalho .cmd dedicado.", C_CYAN),
        ("Isolamento Online", "Ambiente em nuvem multi-empresa com dados segregados por unidade fabril.", C_CYAN_LIGHT),
        ("Cadeados Visuais", "Demarcação vetorial e alertas dinâmicos por faixas de idade da intervenção.", C_ORANGE),
        ("Conformidade Legal", "Alinhamento com NR-10 (Elétrica), NR-12 (Máquinas) e OSHA 1910.147.", C_GREEN)
    ]
    card_w = Inches(2.65)
    card_h = Inches(1.65)
    card_gap = Inches(0.2)
    start_x = Inches(1.2)

    for i, (p_tit, p_desc, p_col) in enumerate(pilares):
        x = start_x + i * (card_w + card_gap)
        c_shape = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(4.3), card_w, card_h)
        c_shape.fill.solid()
        c_shape.fill.fore_color.rgb = RGBColor(15, 23, 42)
        c_shape.line.color.rgb = p_col
        c_shape.line.width = Pt(1.5)
        tf_c = c_shape.text_frame
        tf_c.word_wrap = True
        tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = Inches(0.18)

        p1 = tf_c.paragraphs[0]
        p1.text = p_tit
        p1.font.name = 'Segoe UI'
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = C_WHITE

        p2 = tf_c.add_paragraph()
        p2.text = p_desc
        p2.font.name = 'Segoe UI'
        p2.font.size = Pt(10)
        p2.font.color.rgb = RGBColor(148, 163, 184)
        p2.space_before = Pt(6)

    # Rodapé com Créditos Oficiais
    tb_cred = slide1.shapes.add_textbox(Inches(1.2), Inches(6.3), Inches(11.0), Inches(0.8))
    tf_cr = tb_cred.text_frame
    p_cr1 = tf_cr.paragraphs[0]
    p_cr1.text = "CRÉDITOS DO PROJETO:"
    p_cr1.font.name = 'Segoe UI'
    p_cr1.font.size = Pt(9.5)
    p_cr1.font.bold = True
    p_cr1.font.color.rgb = C_CYAN_LIGHT

    p_cr2 = tf_cr.add_paragraph()
    p_cr2.text = "• Idealizado por: Hadmilton de Melos Cordeiro — Coordenador de Manutenção Elétrica / Supervisor de Estágio\n• Desenvolvimento e implementação: Rodrigo Lopes Chaves de Aguiar — Estagiário de Engenharia de Software"
    p_cr2.font.name = 'Segoe UI'
    p_cr2.font.size = Pt(10)
    p_cr2.font.color.rgb = RGBColor(203, 213, 225)
    p_cr2.space_before = Pt(2)

    # =========================================================================
    # FUNÇÃO HELPER: SLIDE DE CONTEÚDO COM CAPTURA DE TELA EMBUTIDA E CARD RETANGULAR
    # =========================================================================
    def build_content_slide(page_num, tag, title, objetivo, passos, dica, screen_label, screen_obs, img_rel_path):
        slide = prs.slides.add_slide(blank_layout)
        add_bg(slide, C_PAGE_BG)
        add_header(slide, tag, title, page_num)

        # ----------------- COLUNA DA ESQUERDA (INSTRUÇÕES E PROCEDIMENTO) -----------------
        left_w = Inches(5.9)
        left_h = Inches(5.8)
        left_x = Inches(0.75)
        left_y = Inches(1.35)

        # Card de Fundo da Esquerda (Retângulo sem arredondamento excessivo)
        card_l = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, left_y, left_w, left_h)
        card_l.fill.solid()
        card_l.fill.fore_color.rgb = C_CARD_BG
        card_l.line.color.rgb = C_CARD_BORDER
        card_l.line.width = Pt(1)

        # Conteúdo de Texto
        tb_content = slide.shapes.add_textbox(left_x + Inches(0.25), left_y + Inches(0.2), left_w - Inches(0.5), left_h - Inches(1.35))
        tf = tb_content.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        # Seção Objetivo
        p_obj_h = tf.paragraphs[0]
        p_obj_h.text = "🎯 OBJETIVO DESTE PASSO"
        p_obj_h.font.name = 'Segoe UI'
        p_obj_h.font.size = Pt(11)
        p_obj_h.font.bold = True
        p_obj_h.font.color.rgb = C_CYAN

        p_obj = tf.add_paragraph()
        p_obj.text = objetivo
        p_obj.font.name = 'Segoe UI'
        p_obj.font.size = Pt(10.5)
        p_obj.font.color.rgb = C_TEXT_BODY
        p_obj.space_before = Pt(2)
        p_obj.space_after = Pt(8)

        # Seção Passo a Passo
        p_passo_h = tf.add_paragraph()
        p_passo_h.text = "📋 PROCEDIMENTO OPERACIONAL"
        p_passo_h.font.name = 'Segoe UI'
        p_passo_h.font.size = Pt(11)
        p_passo_h.font.bold = True
        p_passo_h.font.color.rgb = C_HEADER_BG
        p_passo_h.space_after = Pt(3)

        for i, item in enumerate(passos, start=1):
            p_item = tf.add_paragraph()
            p_item.text = f"{i}. {item}"
            p_item.font.name = 'Segoe UI'
            p_item.font.size = Pt(10)
            p_item.font.color.rgb = C_TEXT_BODY
            p_item.space_before = Pt(2.5)

        # Caixa de Dica / Atenção no rodapé do card esquerdo (Retangular com detalhe azul)
        dica_box_y = left_y + left_h - Inches(1.15)
        dica_box_h = Inches(0.95)
        dica_box = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_x + Inches(0.25),
            dica_box_y,
            left_w - Inches(0.5),
            dica_box_h
        )
        dica_box.fill.solid()
        dica_box.fill.fore_color.rgb = C_DICA_BG
        dica_box.line.color.rgb = C_DICA_BORDER
        dica_box.line.width = Pt(1)

        # Friso lateral na caixa de dica para acabamento editorial
        dica_stripe = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            left_x + Inches(0.25),
            dica_box_y,
            Inches(0.06),
            dica_box_h
        )
        dica_stripe.fill.solid()
        dica_stripe.fill.fore_color.rgb = C_CYAN
        dica_stripe.line.fill.background()

        tf_d = dica_box.text_frame
        tf_d.word_wrap = True
        tf_d.margin_left = Inches(0.18)
        tf_d.margin_right = Inches(0.12)
        tf_d.margin_top = Inches(0.10)
        
        p_dh = tf_d.paragraphs[0]
        p_dh.text = "💡 RECOMENDAÇÃO TÉCNICA / SEGURANÇA:"
        p_dh.font.name = 'Segoe UI'
        p_dh.font.size = Pt(9.5)
        p_dh.font.bold = True
        p_dh.font.color.rgb = C_CYAN

        p_dt = tf_d.add_paragraph()
        p_dt.text = dica
        p_dt.font.name = 'Segoe UI'
        p_dt.font.size = Pt(9.5)
        p_dt.font.color.rgb = RGBColor(12, 74, 110)
        p_dt.space_before = Pt(2)

        # ----------------- COLUNA DA DIREITA (CAPTURA DE TELA EMBUTIDA) -----------------
        right_w = Inches(5.9)
        right_h = Inches(5.8)
        right_x = Inches(6.8)
        right_y = Inches(1.35)

        # Card de Fundo da Direita (Retângulo limpo e técnico)
        card_r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, right_y, right_w, right_h)
        card_r.fill.solid()
        card_r.fill.fore_color.rgb = C_CARD_BG
        card_r.line.color.rgb = C_CARD_BORDER
        card_r.line.width = Pt(1)

        # Cabeçalho da Área de Captura
        ph_head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, right_y, right_w, Inches(0.42))
        ph_head.fill.solid()
        ph_head.fill.fore_color.rgb = C_HEADER_BG
        ph_head.line.fill.background()
        tf_phh = ph_head.text_frame
        p_phh = tf_phh.paragraphs[0]
        p_phh.text = f"📸 TELA OPERACIONAL: {screen_label.upper()}"
        p_phh.font.name = 'Segoe UI'
        p_phh.font.size = Pt(9.5)
        p_phh.font.bold = True
        p_phh.alignment = PP_ALIGN.CENTER
        p_phh.font.color.rgb = C_WHITE

        # Área de Exibição da Imagem (Moldura técnica de fundo)
        img_container_x = right_x + Inches(0.18)
        img_container_y = right_y + Inches(0.52)
        img_container_w = right_w - Inches(0.36)
        img_container_h = Inches(3.65)

        img_frame = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            img_container_x,
            img_container_y,
            img_container_w,
            img_container_h
        )
        img_frame.fill.solid()
        img_frame.fill.fore_color.rgb = RGBColor(248, 250, 252)
        img_frame.line.color.rgb = RGBColor(226, 232, 240)
        img_frame.line.width = Pt(1)

        # Inserir a Imagem Real com preservação exata da proporção (Aspect Ratio)
        full_img_path = os.path.abspath(img_rel_path)
        if os.path.exists(full_img_path):
            with Image.open(full_img_path) as im:
                orig_w, orig_h = im.size

            aspect = orig_w / orig_h
            fit_w = img_container_w - Inches(0.08)
            fit_h = fit_w / aspect
            if fit_h > (img_container_h - Inches(0.08)):
                fit_h = img_container_h - Inches(0.08)
                fit_w = fit_h * aspect

            # Centralizar na moldura
            pic_x = img_container_x + (img_container_w - fit_w) / 2
            pic_y = img_container_y + (img_container_h - fit_h) / 2
            slide.shapes.add_picture(full_img_path, pic_x, pic_y, width=fit_w, height=fit_h)

        # Rodapé da Captura: O que observar na tela
        obs_y = right_y + Inches(4.28)
        tb_obs = slide.shapes.add_textbox(right_x + Inches(0.25), obs_y, right_w - Inches(0.5), Inches(1.4))
        tf_obs = tb_obs.text_frame
        tf_obs.word_wrap = True
        tf_obs.margin_left = tf_obs.margin_right = tf_obs.margin_top = tf_obs.margin_bottom = 0

        p_obsh = tf_obs.paragraphs[0]
        p_obsh.text = "🔍 ELEMENTOS-CHAVE A DESTACAR NESTA TELA:"
        p_obsh.font.name = 'Segoe UI'
        p_obsh.font.size = Pt(9.5)
        p_obsh.font.bold = True
        p_obsh.font.color.rgb = C_HEADER_BG
        p_obsh.space_after = Pt(2)

        for obs in screen_obs:
            p_o = tf_obs.add_paragraph()
            p_o.text = f"• {obs}"
            p_o.font.name = 'Segoe UI'
            p_o.font.size = Pt(9.0)
            p_o.font.color.rgb = C_TEXT_BODY
            p_o.space_before = Pt(2)

    # =========================================================================
    # DEFINIÇÃO DOS 15 SLIDES DE CONTEÚDO (COM AS CAPTURAS REAIS MAPEADAS)
    # =========================================================================

    slides_data = [
        # SLIDE 2: PASSO 1
        (
            2,
            "MÓDULO 01 • INICIALIZAÇÃO DO SISTEMA",
            "Passo 1: Primeiro Acesso e Conta de Contingência",
            "Inicializar o sistema pela primeira vez cadastrando o administrador principal e assegurar o provisionamento automático da conta de segurança.",
            [
                "Execute o atalho oficial 'EasyLOTOTO for Knauf' na Área de Trabalho ou o arquivo 'INICIAR_EASYLOTOTO.cmd'.",
                "O navegador abrirá automaticamente em 'http://localhost:3001' exibindo o formulário de 'Primeiro acesso'.",
                "Preencha o Nome Completo do responsável pela fábrica.",
                "Defina o Usuário Administrador (ex: 'admin' ou seu login padrão).",
                "Cadastre a Senha mestre com no mínimo 8 caracteres e confirme.",
                "Clique em 'Criar administrador' para liberar o acesso pleno."
            ],
            "O sistema gera em paralelo uma conta oculta ('recovery' / '1234Abcd') permanentemente protegida, para recuperação de emergência caso a senha master seja esquecida.",
            "Tela de Primeiro Acesso (/login)",
            [
                "Card posicionado à direita sobre a imagem de fundo industrial fixa.",
                "Campos de Nome, Usuário, Senha e Confirmação de Senha.",
                "Links inferiores de acesso público: 'Sobre o EasyLOTOTO' e 'Manual de Uso'."
            ],
            "scratch/screenshots/slide_02_primeiro_acesso.png"
        ),

        # SLIDE 3: PASSO 2
        (
            3,
            "MÓDULO 02 • MULTI-TENANT (ONLINE)",
            "Passo 2: Gestão de Empresas e Isolamento de Dados",
            "No ambiente online, cadastrar as filiais ou unidades fabris com isolamento estrito de dados, configurações, plantas e registros.",
            [
                "Acesse o sistema com o perfil 'Global Admin' (administrador corporativo).",
                "No menu superior, selecione a guia 'Empresas & Contas'.",
                "Clique no botão 'Nova Empresa'.",
                "Preencha o Nome da Empresa (ex: 'Knauf Fábrica Camaçari') e o Slug de identificação na URL (ex: 'camacari').",
                "Para alternar entre ambientes, clique em 'Entrar como Admin nesta Empresa'.",
                "O topo exibirá o badge sinalizador: '[Ambiente: Knauf Camaçari]'."
            ],
            "Cada empresa opera sobre uma base de dados SQLite dedicada e segura. Usuários de uma filial não visualizam plantas nem intervenções de outra unidade.",
            "Painel de Empresas e Contas (/companies)",
            [
                "Tabela de empresas cadastradas com status Ativo / Inativo e Slug.",
                "Badge sinalizador do ambiente de trabalho atualmente em operação.",
                "Botões de ação: Entrar como Admin, Editar dados e Desativar filial."
            ],
            "scratch/screenshots/slide_03_empresas_contas.png"
        ),

        # SLIDE 4: PASSO 3
        (
            4,
            "MÓDULO 03 • CONTROLE DE ACESSO",
            "Passo 3: Gestão de Usuários e Perfis de Permissão",
            "Cadastrar a equipe de manutenção e segurança do trabalho concedendo o nível estrito de permissão adequado à função de cada colaborador.",
            [
                "Acesse o menu 'Usuários' através da barra de navegação.",
                "Clique em 'Novo Usuário' para abrir a janela modal de cadastro.",
                "Informe o Nome Completo e o Login de acesso do colaborador.",
                "Defina a Senha de acesso (mínimo de 8 caracteres).",
                "Selecione o Perfil correto: Visualizador, Operador ou Administrador.",
                "Clique em 'Salvar'. O usuário já poderá efetuar login imediatamente."
            ],
            "Hierarquia de Perfis: Visualizador (consulta mapas e relatórios), Operador (registra e desbloqueia etiquetas), Administrador (cria plantas, áreas e regras de KPIs).",
            "Gestão de Usuários e Perfis (/users)",
            [
                "Listagem de usuários ativos com crachá/badge visual de cada perfil.",
                "Modal de criação e edição rápida com redefinição segura de senha.",
                "Proteção contra força bruta: CAPTCHA automático após 3 erros consecutivos."
            ],
            "scratch/screenshots/slide_04_gestao_usuarios.png"
        ),

        # SLIDE 5: PASSO 4
        (
            5,
            "MÓDULO 04 • GESTÃO VISUAL DE PLANTAS",
            "Passo 4: Upload e Estruturação das Plantas Fabris",
            "Inserir as imagens arquitetônicas e diagramas da fábrica que servirão como base para o georreferenciamento dos bloqueios de manutenção.",
            [
                "Acesse o menu 'Plantas e áreas' no menu superior administrativo.",
                "Na barra lateral esquerda, clique no botão '+' ao lado de 'Plantas'.",
                "Envie o arquivo gráfico da fábrica (formatos suportados: SVG, PNG ou JPG).",
                "Nomeie a planta (ex: 'Visão Geral da Fábrica', 'Moinho de Gesso').",
                "Marque como 'Planta raiz' se ela for a tela inicial principal.",
                "Defina plantas filhas vinculadas a áreas para criar zoom multinível."
            ],
            "Prefira arquivos no formato SVG (vetorial). Eles garantem zoom infinito sem perda de nitidez mesmo em telas industriais ultrawide ou tablets de campo.",
            "Editor de Plantas e Áreas (/admin)",
            [
                "Barra lateral esquerda com a lista hierárquica das plantas industriais.",
                "Canvas interativo central com controles de zoom (+/-) e centralização.",
                "Indicação visual da planta raiz (ponto de partida da navegação fabril)."
            ],
            "scratch/screenshots/slide_05_upload_plantas.png"
        ),

        # SLIDE 6: PASSO 5
        (
            6,
            "MÓDULO 05 • DEMARCAÇÃO DE EQUIPAMENTOS",
            "Passo 5: Desenho Vetorial de Áreas e Máquinas",
            "Demarcar com polígonos ou retângulos o perímetro exato de cada máquina, linha de produção ou painel elétrico sobre a planta fabril.",
            [
                "Selecione a planta desejada no painel esquerdo do editor.",
                "Na barra de ferramentas, clique na ferramenta 'Polígono' ou 'Retângulo'.",
                "Clique no mapa para traçar os pontos contornando o equipamento.",
                "Dê um duplo clique no último ponto para fechar a figura demarcada.",
                "No inspetor lateral direito, preencha o Nome (ex: 'Ensacadeira 02').",
                "Preencha o 'Código': este é o tag de integração externa (ex: 'ENS-02').",
                "Ajuste a cor da demarcação e a transparência desejada."
            ],
            "O campo 'Código' é fundamental: é através dele que o sistema correlaciona ordens de serviço do SAP, TOTVS ou planilhas Excel diretamente ao mapa.",
            "Demarcação Vetorial no Canvas (/admin)",
            [
                "Vértices editáveis do polígono demarcado sobre o layout da máquina.",
                "Inspetor lateral com Nome, Código, Opacidade e Cor padrão da área.",
                "Ferramentas de exclusão, duplicação e ordenação de camadas visuais."
            ],
            "scratch/screenshots/slide_06_demarcacao_vetorial.png"
        ),

        # SLIDE 7: PASSO 6
        (
            7,
            "MÓDULO 06 • BLOQUEIOS LOTOTO",
            "Passo 6: Parametrização do Cadeado LOTOTO na Área",
            "Configurar a regra lógica que aciona a exibição do cadeado visual sobre a máquina e personalizar os dados do balão informativo de inspeção.",
            [
                "Selecione a área demarcada no editor de plantas.",
                "No inspetor direito, localize a seção 'Cadeado (LOTOTO)'.",
                "Em 'Filtro do cadeado', escreva a fórmula de ativação lógica.",
                "Exemplo clássico: [Equipamento] = \"Moinho 1\" AND [Status] = \"Bloqueado\"",
                "Em 'Informações mostradas ao clicar no cadeado', configure os textos.",
                "Exemplo: Ordem: {OS} | Responsável: {Mecânico} | Desde: {Data_Bloqueio}",
                "Clique em 'Salvar área'. O cadeado responderá instantaneamente."
            ],
            "Ao haver uma intervenção ativa correspondente, o cadeado surge na cor exata da sua idade de bloqueio, com um contador numérico do total de bloqueios no local.",
            "Configuração do Cadeado LOTOTO (/admin)",
            [
                "Campo 'Filtro do cadeado' com sintaxe de colchetes validada em tempo real.",
                "Campo de interpolação textual com chaves {Campo} para o balão popup.",
                "Pré-visualização do badge dinâmico do cadeado sobre a área da máquina."
            ],
            "scratch/screenshots/slide_07_parametrizacao_cadeado.png"
        ),

        # SLIDE 8: PASSO 7
        (
            8,
            "MÓDULO 07 • MOTOR DE REGRAS",
            "Passo 7: Motor de Expressões, Operadores e Fórmulas",
            "Dominar a elaboração de fórmulas condicionais booleanas para filtros de cadeados e cálculos estatísticos sem necessidade de programação.",
            [
                "Identificadores de campos: sempre entre colchetes, ex: [Setor], [Tipo_Energia].",
                "Valores de texto: sempre entre aspas duplas, ex: \"Elétrica\", \"Aberta\".",
                "Operadores lógicos: utilize AND, OR, NOT e parênteses para agrupar.",
                "Operadores relacionais: =, !=, >, <, >=, <=, CONTAINS, STARTSWITH.",
                "Exemplo combinado: [Setor] = \"Moagem\" AND ([Tipo] = \"Elétrica\" OR [Prioridade] > 3)",
                "Utilize o botão 'Testar expressão' para verificar a sintaxe imediatamente."
            ],
            "Para consultar todas as funções disponíveis, acesse a rota '/manual' com documentação interativa completa no formato Swagger UI.",
            "Editor de Expressões com Autocomplete (/admin)",
            [
                "Sugestão inteligente de nomes de campos cadastrados ao digitar '['.",
                "Coloração sintática destacada (verde para strings, azul para operadores).",
                "Feedback de validação instantâneo (Sintaxe válida em verde / Erro explicativo)."
            ],
            "scratch/screenshots/slide_08_motor_expressoes.png"
        ),

        # SLIDE 9: PASSO 8
        (
            9,
            "MÓDULO 08 • ESTRUTURA DE DADOS",
            "Passo 8: Campos Personalizados e Campos Calculados",
            "Estruturar os campos das ordens de bloqueio da fábrica e criar colunas inteligentes calculadas automaticamente por fórmulas em tempo real.",
            [
                "Acesse o menu 'Campos' na barra superior.",
                "Clique em 'Novo Campo' para definir uma nova coluna do formulário LOTOTO.",
                "Escolha o Tipo: Texto, Número, Data e Hora, Seleção de Opções ou Foto.",
                "Para criar um Campo Calculado, marque 'Este campo é calculado por fórmula'.",
                "Insira a fórmula. Exemplo de cálculo de horas de intervenção ativa:\nHOURS_BETWEEN([Data_Inicio], NOW())",
                "Arraste os campos na tabela para ordenar a exibição nas telas e relatórios."
            ],
            "Campos do tipo 'Foto' permitem anexar fotos das etiquetas físicas, cadeados instalados e dispositivos de bloqueio aplicados, com visualização em Lightbox com zoom.",
            "Tabela de Campos Personalizados (/fields)",
            [
                "Listagem com Rótulo, Chave interna, Tipo e Sinalizador de Calculado.",
                "Botões de reordenação vertical rápida e edição de propriedades.",
                "Campos de fotos de evidência com suporte a múltiplos anexos por ordem."
            ],
            "scratch/screenshots/slide_09_campos_personalizados.png"
        ),

        # SLIDE 10: PASSO 9
        (
            10,
            "MÓDULO 09 • GESTÃO À VISTA",
            "Passo 9: Criação de Painéis e KPIs de Segurança",
            "Construir indicadores gerenciais em tempo real exibidos diretamente sobre o mapa da fábrica para acompanhamento visual de líderes e operadores.",
            [
                "Acesse o menu 'KPIs' na barra administrativa.",
                "Clique em 'Novo KPI' para configurar um indicador visual.",
                "Defina o Título (ex: 'Bloqueios Ativos') e o Subtítulo explicativo.",
                "Escolha o Formato: 'Card de Valor Único' ou 'Barra Proporcional'.",
                "Insira a fórmula agregadora. Exemplos úteis:\n• Contar bloqueios abertos: COUNT_IF([Status] = \"Ativo\")\n• Contar bloqueios críticos: COUNT_IF(HOURS_BETWEEN([Data], NOW()) > 72)",
                "Defina faixas de metas com cores de alerta (Verde, Amarelo e Vermelho)."
            ],
            "Os KPIs operam com atualização reativa em tempo real. Sempre que um operador cadastra, edita ou desbloqueia um cadeado, os valores recalculam sem recarregar a tela.",
            "Cards de KPIs no Topo do Mapa (/)",
            [
                "Painel superior horizontal integrado com cards de métricas em destaque.",
                "Barras proporcionais comparativas por tipo de energia perigosa ou setor.",
                "Sinalizadores de conformidade com mudança de cor por faixas de risco."
            ],
            "scratch/screenshots/slide_10_kpis_paineis.png"
        ),

        # SLIDE 11: PASSO 10
        (
            11,
            "MÓDULO 10 • TEMPO E REINCIDÊNCIA",
            "Passo 10: Critérios de Idade do Bloqueio e Reincidência",
            "Configurar os limites de tempo das intervenções para mudança automática de cor dos cadeados e monitorar reincidências de falhas na mesma máquina.",
            [
                "Acesse o menu 'Configurações'.",
                "Na seção 'Cor do cadeado pela idade', visualize as faixas configuradas:",
                "• Recente (a partir de 0h): Cadeado Azul.",
                "• Mais de 24h: Cadeado Amarelo de Atenção.",
                "• Mais de 72h: Cadeado Laranja Crítico.",
                "• Mais de 7 dias (168h): Cadeado Vermelho Grave.",
                "Clique em 'Adicionar faixa' ou edite as horas e cores conforme a planta.",
                "Em 'Reincidência', defina o prazo (ex: 7 dias) para flag de reincidência."
            ],
            "O cadeado no mapa adota a cor da intervenção mais antiga daquela área demarcada. Isso garante que intervenções esquecidas fiquem imediatamente visíveis para a chefia.",
            "Configuração de Cores e Reincidência (/settings)",
            [
                "Tabela de faixas com Seletor de Cor, Horas Mínimas e Rótulo da legenda.",
                "Controle numérico de dias para apuração da taxa de reincidência de falha.",
                "Legenda dinâmica de cores exibida no canto inferior do mapa operacional."
            ],
            "scratch/screenshots/slide_11_criterios_idade.png"
        ),

        # SLIDE 12: PASSO 11
        (
            12,
            "MÓDULO 11 • INTEGRAÇÃO DE DADOS",
            "Passo 11: Importação de Dados Externos (SAP, Excel, Poka)",
            "Carregar planilhas de ordens de manutenção provenientes de sistemas ERP/GMAO para atualizar visualmente os bloqueios de forma massiva e rápida.",
            [
                "Acesse o menu 'Registros' e clique no botão 'Importar Excel'.",
                "Arraste ou selecione a planilha de ordens (formatos .xlsx ou .xlsm).",
                "O sistema analisa os cabeçalhos e sugere o vínculo com os campos locais.",
                "Ajuste manualmente o mapeamento de colunas caso necessário.",
                "Verifique a tela de pré-visualização (linhas prontas vs dados inconsistentes).",
                "Clique em 'Salvar como perfil' (ex: 'Padrão SAP Manutenção').",
                "Confirme a importação. Os registros e cadeados aparecem de imediato."
            ],
            "O sistema calcula uma assinatura de integridade (hash) para cada linha da planilha, impedindo duplicações indesejadas caso o mesmo arquivo seja reimportado.",
            "Assistente de Importação com Mapeamento (/import)",
            [
                "Mapeamento visual de colunas da planilha externa para campos do LOTOTO.",
                "Seletor de perfis de importação salvos para execução rápida em 1 clique.",
                "Tabela de auditoria com resumo de linhas importadas e histórico de validação."
            ],
            "scratch/screenshots/slide_12_importacao_excel.png"
        ),

        # SLIDE 13: PASSO 12
        (
            13,
            "MÓDULO 12 • SEGURANÇA E BACKUPS",
            "Passo 12: Backups Locais, Rotinas Automáticas e Área de Risco",
            "Garantir a preservação dos dados da fábrica por meio de cópias de segurança compactadas e gerenciar exclusões controladas por palavra-chave.",
            [
                "Acesse o menu 'Backup' na barra de ferramentas administrativa.",
                "Na aba 'Backups locais', escolha 'Somente dados' ou 'Dados + configurações'.",
                "Clique em 'Criar e baixar backup' para gerar o arquivo compactado .zip.",
                "Em caso de troca de computador, use 'Restaurar backup' e selecione o arquivo.",
                "Na aba 'Conta', localize a 'Área de Risco' para manutenções destrutivas.",
                "Para formatar o sistema online, digite 'DELETE.ALL' na confirmação."
            ],
            "O servidor executa automaticamente um backup diário rotativo mantendo sempre os arquivos mais recentes no disco, garantindo tolerância a falhas sem intervenção humana.",
            "Painel de Backups e Área de Risco (/backup)",
            [
                "Lista de backups locais gerados com tamanho, data e botão de restauração.",
                "Aba Conta com área de risco destacada para limpeza de dados e logs.",
                "Janela de confirmação por digitação de palavra-chave de segurança."
            ],
            "scratch/screenshots/slide_13_backups_seguranca.png"
        ),

        # SLIDE 14: PASSO 13
        (
            14,
            "MÓDULO 13 • SINCRONIZAÇÃO HÍBRIDA",
            "Passo 13: Sincronização Dinâmica (Offline ↔ Online)",
            "Transmitir registros de manutenção da fábrica local sem rede para o ambiente corporativo online através de leitura visual contínua por QR Code.",
            [
                "Na máquina transmissora (offline), acesse 'Backup > Sincronização'.",
                "Clique em 'Gerar código de sincronização' (escolha Dados ou Completo).",
                "Clique em 'Ver código gerado' para abrir a transmissão contínua de QR Codes.",
                "No notebook receptor (online), acesse 'Receber sincronização'.",
                "Clique em 'Escanear pela câmera' e aponte a webcam para a tela transmissora.",
                "Os quadros são lidos em alta velocidade (streaming contínuo).",
                "Confirme o resumo das alterações para aplicar a sincronização."
            ],
            "Alternativa manual: No modal de transmissão, você pode clicar em 'Copiar código de sincronização' e colar diretamente na outra máquina via pen-drive ou rede local.",
            "Transmissão Visual por QR Code Dinâmico (/backup)",
            [
                "Visualizador de QR Codes com alternância acelerada e limpa.",
                "Campo de transferência manual com código Base64 gzippado em lote.",
                "Interface do scanner por câmera com barra de progresso dos quadros capturados."
            ],
            "scratch/screenshots/slide_14_sincronizacao_qrcode.png"
        ),

        # SLIDE 15: PASSO 14
        (
            15,
            "MÓDULO 14 • CONSULTA E IA",
            "Passo 14: Manual Swagger Interativo e Assistência por IA",
            "Consultar a biblioteca de funções e expressões do sistema através de uma interface estilo Swagger e utilizar a skill otimizada para Inteligência Artificial.",
            [
                "Acesse a rota pública '/manual' através do link na tela de login.",
                "A interface organiza os recursos em tags: CADEADO, KPI, CAMPO, DATA e LÓGICA.",
                "Clique em qualquer função para ver sintaxe, parâmetros, retorno e exemplos.",
                "Utilize o botão 'Copiar' para colar fórmulas prontas direto no editor.",
                "No topo, clique em 'Baixar Versão para IA' para salvar o 'easylototo-skill.md'.",
                "Anexe esse arquivo no ChatGPT, Claude ou Gemini para obter suporte técnico."
            ],
            "A skill ensina a IA todas as regras do EasyLOTOTO, capacitando-a a escrever filtros complexos, resolver dúvidas de operadores e sugerir cálculos de KPIs da fábrica.",
            "Documentação Interativa estilo Swagger (/manual)",
            [
                "Layout clássico Swagger UI com badges coloridos (GET, POST, EXPR).",
                "Botão de download de alta visibilidade da skill especializada para IA.",
                "Botão para baixar o Manual Técnico oficial completo em PDF."
            ],
            "scratch/screenshots/slide_15_manual_swagger_ia.png"
        ),

        # SLIDE 16: ENCERRAMENTO E BOAS PRÁTICAS
        (
            16,
            "MÓDULO 15 • ROTINA OPERACIONAL",
            "Passo 15: Checklist Diário e Boas Práticas LOTOTO",
            "Estabelecer a rotina diária padronizada para os técnicos de manutenção e segurança garantindo tolerância zero a acidentes com energia perigosa.",
            [
                "Abertura de Turno: Inspecionar o mapa em busca de cadeados amarelos, laranjas ou vermelhos para identificar intervenções prolongadas.",
                "Início do Bloqueio: Antes de descer para a área, registrar o bloqueio com Tag da máquina, número da OS, mecânico responsável e foto da trava física.",
                "Teste de Energia Zero: Testar o botão de partida local (Tryout) para garantir energia zero antes de iniciar qualquer intervenção mecânica ou elétrica.",
                "Conclusão da Ordem: Ao retirar o cadeado físico, registrar o desbloqueio no sistema. O cadeado desaparecerá automaticamente da planta fabril.",
                "Fechamento Semanal: Gerar backup preventivo (.zip) e salvar em disco externo."
            ],
            "NR-10 / NR-12: Nenhum colaborador deve intervir em máquinas sem bloqueio físico individual e sinalização inequívoca no mapa da fábrica.",
            "Mapa da Fábrica Operacional em Produção (/)",
            [
                "Visão geral dos cadeados ativos distribuídos pelos setores fabris.",
                "Painel superior com status dos indicadores de segurança atualizados.",
                "Créditos do projeto no rodapé: Hadmilton de Melos Cordeiro e Rodrigo Lopes Chaves de Aguiar."
            ],
            "scratch/screenshots/slide_16_mapa_operacional.png"
        )
    ]

    for data in slides_data:
        build_content_slide(*data)

    output_path = os.path.abspath("Passo_a_Passo_EasyLOTOTO_for_Knauf.pptx")
    prs.save(output_path)
    print(f"Apresentação gerada com sucesso em: {output_path}")

if __name__ == "__main__":
    create_presentation()
