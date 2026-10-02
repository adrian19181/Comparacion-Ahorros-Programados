# ==============================================================================
    # 4. GRÁFICOS DINÁMICOS
    # ==============================================================================
    st.divider()
    st.subheader("📊 Análisis Gráfico de Ahorros")

    import plotly.graph_objects as go
    from plotly.subplots import make_subplots

    # Función auxiliar para convertir las cadenas formateadas a números puros para graficar
    def str_to_float(val_str):
        if val_str in ("-", "", None): return 0.0
        clean_str = str(val_str).replace("$", "").replace("%", "").replace(".", "").replace(",", ".")
        try:
            return float(clean_str)
        except:
            return 0.0

    years = list(datos.keys())
    
    cap_depositado = [str_to_float(datos[y]["dep_sin_int"]) for y in years]
    int_ganado = [str_to_float(datos[y]["int_ganados"]) for y in years]
    
    tna_vals = [str_to_float(datos[y]["tna"]) for y in years]
    tea_vals = [str_to_float(datos[y]["tea"]) for y in years]
    tir_vals = [str_to_float(datos[y]["tir"]) for y in years]
    
    int_mensual_pond = [str_to_float(datos[y]["int_mensual_pond"]) for y in years]
    int_diario_prom = [str_to_float(datos[y]["int_diario_prom"]) for y in years]

    tab_comp, tab_tasas, tab_int = st.tabs(["💰 Capital vs Intereses", "📉 Comparativa de Tasas", "📈 Interés Promedio"])
    
    # CONFIGURACIÓN ESTRICTA PARA MÓVIL: Bloquea todo tipo de zoom, paneo y toque en la pantalla
    plotly_config = {
        'displayModeBar': False,
        'staticPlot': True  # Convierte el gráfico en una imagen estática (ideal para táctil)
    }

    with tab_comp:
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=years, y=cap_depositado, name="Capital Depositado", marker_color="#1E3A8A", text=[f"${v:,.0f}" for v in cap_depositado], textposition="inside"))
        fig1.add_trace(go.Bar(x=years, y=int_ganado, name="Intereses Ganados", marker_color="#10B981", text=[f"${v:,.0f}" for v in int_ganado], textposition="inside"))
        
        fig1.update_layout(
            barmode='stack',
            title="Composición del Ahorro Final",
            template="plotly_dark",
            yaxis_tickprefix="$",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            dragmode=False,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        fig1.update_xaxes(fixedrange=True)
        fig1.update_yaxes(fixedrange=True)
        st.plotly_chart(fig1, use_container_width=True, config=plotly_config)

    with tab_tasas:
        fig2 = go.Figure()
        # Colores completamente distintos para que no se repitan visualmente
        fig2.add_trace(go.Bar(x=years, y=tna_vals, name="TNA", marker_color="#38BDF8", text=[f"{v:.2f}%" for v in tna_vals], textposition="auto"))
        fig2.add_trace(go.Bar(x=years, y=tea_vals, name="TEA", marker_color="#F43F5E", text=[f"{v:.2f}%" for v in tea_vals], textposition="auto"))
        fig2.add_trace(go.Bar(x=years, y=tir_vals, name="TIR", marker_color="#FBBF24", text=[f"{v:.2f}%" for v in tir_vals], textposition="auto"))
        
        fig2.update_layout(
            barmode='group',
            title="Evolución de Tasas de Rendimiento",
            template="plotly_dark",
            yaxis_ticksuffix="%",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            dragmode=False,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        fig2.update_xaxes(fixedrange=True)
        fig2.update_yaxes(fixedrange=True)
        st.plotly_chart(fig2, use_container_width=True, config=plotly_config)

    with tab_int:
        # Usamos doble eje Y para que la línea de interés diario (pequeña) no se aplaste contra la mensual
        fig3 = make_subplots(specs=[[{"secondary_y": True}]])
        
        fig3.add_trace(go.Scatter(
            x=years, y=int_mensual_pond, mode='lines+markers+text', name="Int. Mensual", 
            line=dict(color="#00D1B2", width=4),
            marker=dict(size=12, color="#00D1B2"),
            text=[f"${v:,.2f}" for v in int_mensual_pond], 
            textposition="top center",
            textfont=dict(size=13, color="#00D1B2")
        ), secondary_y=False)

        fig3.add_trace(go.Scatter(
            x=years, y=int_diario_prom, mode='lines+markers+text', name="Int. Diario", 
            line=dict(color="#A855F7", width=4, dash='dot'),
            marker=dict(size=12, color="#A855F7"),
            text=[f"${v:,.2f}" for v in int_diario_prom], 
            textposition="bottom center",
            textfont=dict(size=13, color="#A855F7")
        ), secondary_y=True)
        
        fig3.update_layout(
            title="Interés Mensual vs Diario",
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            dragmode=False,
            margin=dict(l=10, r=10, t=50, b=10)
        )
        # Ocultar las grillas extra y fijar rangos para que los textos no se corten por arriba/abajo
        fig3.update_xaxes(fixedrange=True, showgrid=False)
        fig3.update_yaxes(fixedrange=True, secondary_y=False, showgrid=False, range=[0, max(int_mensual_pond) * 1.3])
        fig3.update_yaxes(fixedrange=True, secondary_y=True, showgrid=False, range=[0, max(int_diario_prom) * 1.3], showticklabels=False)

        st.plotly_chart(fig3, use_container_width=True, config=plotly_config)