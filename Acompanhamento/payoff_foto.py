from manim import *
import numpy as np
import random

# ================================================================
# CONFIGURAÇÃO
# ================================================================

config.pixel_width = 1080
config.pixel_height = 1920

class PayoffConexao3D(ThreeDScene):

    def construct(self):

        # ============================================================
        # PARÂMETROS E TRACKERS
        # ============================================================

        # Trackers permitem que os valores sejam animados suavemente
        tracker_K = ValueTracker(100)
        tracker_T = ValueTracker(5)

        ST = 125

        tempos = np.array([0, 1, 2, 3, 4, 5])
        precos = np.array([90, 95, 88, 105, 115, ST])

        # Função para interpolar o preço, permitindo que o ponto deslize pela linha
        def get_preco(t):
            return np.interp(t, tempos, precos)

        # ============================================================
        # CÂMERA
        # ============================================================

        self.set_camera_orientation(
            phi=0 * DEGREES,
            theta=-90 * DEGREES,
            zoom=1.15
        )

        # ============================================================
        # TÍTULO
        # ============================================================

        titulo = Text(
            'O payoff como uma "foto" do preço da ação',
            font_size=39
        )

        titulo.to_edge(UP, buff=0.30)
        self.add_fixed_in_frame_mobjects(titulo)
        self.play(Write(titulo))

        # ============================================================
        # INFORMAÇÕES DO ESTADO (COM ATUALIZAÇÃO DINÂMICA)
        # ============================================================

        strike_text = always_redraw(lambda: Text(
            f"Strike: K = {tracker_K.get_value():.0f}",
            font_size=27,
            color=RED
        ).move_to(UP * 5.75 + LEFT * 2.6))

        estado_text = always_redraw(lambda: Text(
            f"Preço no vencimento: ST = {get_preco(tracker_T.get_value()):.0f}",
            font_size=27,
            color=BLUE
        ).move_to(UP * 5.75 + RIGHT * 2.2))

        self.add_fixed_in_frame_mobjects(strike_text, estado_text)
        self.play(Write(strike_text), Write(estado_text))

        # ============================================================
        # EIXOS 3D
        # ============================================================

        axes = ThreeDAxes(
            x_range=[0, 5.5, 1],
            y_range=[50, 150, 10],
            z_range=[0, 50, 10],
            x_length=6.5,
            y_length=9.0,
            z_length=5.0,
            axis_config={"include_numbers": False}
        )

        axes.shift(DOWN * 1.25)
        self.play(Create(axes))

        # ============================================================
        # NOMES DOS EIXOS
        # ============================================================

        nome_x = Text("Tempo (t)", font_size=24).next_to(axes.x_axis.get_end(), DOWN, buff=0.15)
        nome_y = Text("Preço da ação (S)", font_size=24).next_to(axes.y_axis.get_end(), LEFT, buff=0.15)
        nome_z = Text("Payoff (Z)", font_size=24).next_to(axes.z_axis.get_end(), RIGHT, buff=0.15)

        self.play(Write(nome_x), Write(nome_y), Write(nome_z))

        # ============================================================
        # STRIKE (COM ATUALIZAÇÃO DINÂMICA)
        # ============================================================

        strike_line = always_redraw(lambda: axes.plot_parametric_curve(
            lambda t: np.array([t, tracker_K.get_value(), 0]),
            t_range=[0, tracker_T.get_value()],
            color=RED
        ))

        self.play(Create(strike_line))

        # ============================================================
        # TRAJETÓRIA PRINCIPAL
        # ============================================================

        pontos_xy = [axes.c2p(t, p, 0) for t, p in zip(tempos, precos)]

        caminho_principal = VMobject()
        caminho_principal.set_points_as_corners(pontos_xy)
        caminho_principal.set_stroke(color=BLUE, width=6)

        ponto_ST = always_redraw(lambda: Dot3D(
            axes.c2p(tracker_T.get_value(), get_preco(tracker_T.get_value()), 0),
            color=BLUE,
            radius=0.14
        ))

        self.play(Create(caminho_principal), run_time=2)
        self.play(FadeIn(ponto_ST))

        # ============================================================
        # FOTO
        # ============================================================

        foto = SurroundingRectangle(ponto_ST, color=WHITE, buff=0.25)
        texto_foto = Text("FOTO", font_size=28).move_to(UP * 5.15)
        self.add_fixed_in_frame_mobjects(texto_foto)

        flash = Rectangle(width=9, height=3.0, stroke_width=8, stroke_opacity=0)
        self.add_fixed_in_frame_mobjects(flash)

        self.play(flash.animate.set_stroke(opacity=1), run_time=0.15)
        self.play(flash.animate.set_stroke(opacity=0), run_time=0.35)

        self.play(Create(foto), Write(texto_foto))
        self.wait(1)

        # ============================================================
        # FRASE EXPLICATIVA
        # ============================================================

        explicacao = Text("No vencimento, só importa o estado final ST", font_size=25).move_to(UP * 4.72)
        self.add_fixed_in_frame_mobjects(explicacao)
        self.play(Write(explicacao))
        self.wait(2)

        # LIMPA FOTO
        self.play(FadeOut(texto_foto), FadeOut(explicacao), FadeOut(foto))

        # ============================================================
        # MUDA A CÂMERA
        # ============================================================

        self.move_camera(
            phi=70 * DEGREES,
            theta=-45 * DEGREES,
            zoom=1.15,
            run_time=2.5
        )

        # ============================================================
        # CURVA DE PAYOFF (COM ATUALIZAÇÃO DINÂMICA)
        # ============================================================

        payoff_curve = always_redraw(lambda: axes.plot_parametric_curve(
            lambda s: np.array([
                tracker_T.get_value(),
                s,
                max(s - tracker_K.get_value(), 0)
            ]),
            t_range=[50, 150],
            color=YELLOW
        ))

        self.play(Create(payoff_curve), run_time=2)

        # ============================================================
        # CONEXÃO ST → PAYOFF (COM ATUALIZAÇÃO DINÂMICA)
        # ============================================================
        
        def get_linha_conexao():
            t_val = tracker_T.get_value()
            k_val = tracker_K.get_value()
            st_val = get_preco(t_val)
            payoff_val = max(st_val - k_val, 0)
            
            # Se o payoff for 0, retornamos um objeto vazio pra evitar erros no DashedLine
            if payoff_val <= 0.001:
                return VMobject()
            
            return DashedLine(
                axes.c2p(t_val, st_val, 0),
                axes.c2p(t_val, st_val, payoff_val),
                color=WHITE
            )

        linha_conexao = always_redraw(get_linha_conexao)

        ponto_payoff = always_redraw(lambda: Dot3D(
            axes.c2p(
                tracker_T.get_value(),
                get_preco(tracker_T.get_value()),
                max(get_preco(tracker_T.get_value()) - tracker_K.get_value(), 0)
            ),
            color=YELLOW,
            radius=0.14
        ))

        self.play(Create(linha_conexao), FadeIn(ponto_payoff))

        # ============================================================
        # FÓRMULA E RESULTADOS (COM ATUALIZAÇÃO DINÂMICA)
        # ============================================================

        formula = Text("Payoff = max(ST - K, 0)", font_size=34).move_to(UP * 5.35)
        self.add_fixed_in_frame_mobjects(formula)
        self.play(Write(formula))
        self.wait(1)

        calculo = always_redraw(lambda: Text(
            f"max({get_preco(tracker_T.get_value()):.0f} - {tracker_K.get_value():.0f}, 0) = {max(get_preco(tracker_T.get_value()) - tracker_K.get_value(), 0):.0f}",
            font_size=29
        ).move_to(UP * 4.85))
        
        self.add_fixed_in_frame_mobjects(calculo)
        self.play(Write(calculo))
        self.wait(1.5)

        resultado = always_redraw(lambda: Text(
            f"Payoff = {max(get_preco(tracker_T.get_value()) - tracker_K.get_value(), 0):.0f}",
            font_size=35,
            color=YELLOW
        ).move_to(UP * 4.40))

        self.add_fixed_in_frame_mobjects(resultado)
        self.play(Write(resultado), Indicate(ponto_payoff, scale_factor=2))
        self.wait(2)

        # ============================================================
        # CAMINHOS ALTERNATIVOS
        # ============================================================

        caminhos_alternativos = VGroup()
        for _ in range(15):
            preco_atual = 90
            pontos_alt = [axes.c2p(0, preco_atual, 0)]
            for t in range(1, 6):
                preco_atual += random.uniform(-20, 20)
                preco_atual = np.clip(preco_atual, 55, 145)
                pontos_alt.append(axes.c2p(t, preco_atual, 0))
            
            caminho_alt = VMobject()
            caminho_alt.set_points_as_corners(pontos_alt)
            caminho_alt.set_stroke(color=GREY, width=2, opacity=0.25)
            caminhos_alternativos.add(caminho_alt)

        self.play(FadeIn(caminhos_alternativos), run_time=3)
        self.wait(3)

        # ============================================================
        # ETAPA 2 — VARIAÇÃO ANIMADA DO STRIKE
        # ============================================================

        titulo_strike = Text("E se o Strike K mudar?", font_size=38).move_to(UP * 5.45)
        self.add_fixed_in_frame_mobjects(titulo_strike)

        # Substitui a fórmula genérica pelo novo título
        self.play(FadeOut(formula), Write(titulo_strike))
        self.wait(1)

        # O Manim cuida da animação: basta dizer para onde o K vai!
        self.play(tracker_K.animate.set_value(80), run_time=2)
        self.wait(1)

        self.play(tracker_K.animate.set_value(140), run_time=3)
        self.wait(1)

        self.play(tracker_K.animate.set_value(100), run_time=2)
        self.wait(2)

        self.play(FadeOut(titulo_strike))

        # ============================================================
        # ETAPA 3 — VARIAÇÃO ANIMADA DO TEMPO
        # ============================================================

        titulo_tempo = Text("E se o tempo de vencimento T mudar?", font_size=36).move_to(UP * 5.45)
        self.add_fixed_in_frame_mobjects(titulo_tempo)
        self.play(Write(titulo_tempo))
        self.wait(1)

        # Agora animamos o Tempo deslizando pelo eixo X!
        self.play(tracker_T.animate.set_value(1), run_time=4)
        self.wait(1)

        self.play(tracker_T.animate.set_value(3), run_time=2)
        self.wait(1)

        self.play(tracker_T.animate.set_value(5), run_time=2)
        self.wait(2)

        self.play(FadeOut(titulo_tempo))

        # ============================================================
        # FRASE FINAL
        # ============================================================

        self.play(FadeOut(calculo), FadeOut(resultado))

        frase_final = Text(
            "O payoff depende do estado final, do Strike e do vencimento",
            font_size=27
        ).move_to(UP * 5.15)
        
        self.add_fixed_in_frame_mobjects(frase_final)
        self.play(Write(frase_final))
        
        self.wait(4)