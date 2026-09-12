if doc_quijada and st.button("🔗 Unir y Guardar en Historial", type="primary"):
    try:
        fecha_texto = fecha_estudio.strftime("%d/%m/%Y") if fecha_estudio else None
        merged_bytes, nombre_detectado = unir_informe_renal(
            "plantilla_renal.docx", doc_quijada,
            fecha_estudio=fecha_texto,
            medico_referente=medico_referente,
            cedula=cedula_paciente
        )
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        nombre_elegido = nombre_para_guardar.strip() if nombre_para_guardar.strip() else nombre_detectado
        nombre_base = nombre_elegido.replace(' ', '_') if nombre_elegido else "SinNombre"
        nombre_hist = f"{timestamp}_Renal_{nombre_base}.docx"
        guardar_en_historial(nombre_hist, merged_bytes.getvalue())
        st.success(f"✅ Guardado como '{nombre_hist}'. Lo encuentras más abajo, en Historial de Informes.")

        nombre_descarga = f"Informe Renal - {nombre_elegido}.docx" if nombre_elegido else "Informe Renal.docx"
        st.download_button(
            "📥 Descargar este informe (.docx)",
            data=merged_bytes.getvalue(),
            file_name=nombre_descarga,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        if verificacion_renal_activada:
            advertencias_renal = verificar_informe_renal_basico(merged_bytes.getvalue(), nombre_detectado)
            if advertencias_renal:
                st.warning("🔍 Verificación automática — revisa esto antes de enviarlo:")
                for adv in advertencias_renal:
                    st.markdown(f"- {adv}")
            else:
                st.success("🔍 Verificación automática: no se detectaron problemas. (No reemplaza la revisión clínica.)")
    except Exception as e:
        st.error(f"Error al unir los documentos: {e}")s