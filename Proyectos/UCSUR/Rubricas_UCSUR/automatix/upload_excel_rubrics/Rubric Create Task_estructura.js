// ==UserScript==
// @name         Crear estructura de tareas en Canvas (Versión Mejorada)
// @namespace    https://cientificavirtual.cientifica.edu.pe/
// @version      2.2
// @description  Crea grupos y tareas predefinidas con manejo robusto de errores
// @match        https://cientificavirtual.cientifica.edu.pe/courses/*/assignments
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // Configuración
    const config = {
        estructura: [
            { grupo: "Evaluación Diagnóstica", tarea: "Evaluación Diagnóstica" },
            { grupo: "Evaluación Continua 1", tarea: "Evaluación Continua 1" },
            { grupo: "Evaluación Continua 2", tarea: "Evaluación Continua 2" },
            { grupo: "Evaluación Continua 3", tarea: "Evaluación Continua 3" },
            { grupo: "Examen Parcial", tarea: "Examen Parcial" },
            { grupo: "Examen Final", tarea: "Examen Final" },
        ],
        delayBetweenRequests: 1000, // Tiempo entre solicitudes en ms
        puntosPorTarea: 20,
        grupoProtegido: "Tareas" // Grupo que no se eliminará
    };

    // Estado del script
    const state = {
        courseId: null,
        csrfToken: null,
        isRunning: false
    };

    // Utilidades
    const utils = {
        // Función mejorada para obtener el ID del curso
        getCourseId() {
            const matches = [...window.location.pathname.matchAll(/\/courses\/(\d+)/g)];
            return matches[0]?.[1] || window.ENV?.COURSE_ID || window.ENV?.course_id || null;
        },

        // Función mejorada para obtener el token CSRF
        getCSRFToken() {
            return document.querySelector('meta[name="csrf-token"]')?.content ||
                   document.querySelector('[name="csrf-token"]')?.content ||
                   window.ENV?.csrf_token ||
                   this.getCsrfTokenFromCookies();
        },

        // Obtener CSRF de cookies
        getCsrfTokenFromCookies() {
            const csrfRegex = new RegExp('^_csrf_token=(.*)$');
            const cookies = document.cookie.split(';');
            for (const cookie of cookies) {
                const match = csrfRegex.exec(cookie.trim());
                if (match) {
                    return decodeURIComponent(match[1]);
                }
            }
            return null;
        },

        // Mostrar notificación
        showNotification(message, type = 'info', duration = 5000) {
            const colors = {
                info: '#2D9CDB',
                success: '#6FCF97',
                warning: '#F2C94C',
                error: '#EB5757'
            };

            const existing = document.getElementById('canvas-helper-notification');
            if (existing) existing.remove();

            const notification = document.createElement('div');
            notification.id = 'canvas-helper-notification';
            notification.style = `
                position: fixed;
                top: 20px;
                right: 20px;
                padding: 15px;
                background: ${colors[type]};
                color: white;
                border-radius: 5px;
                box-shadow: 0 3px 10px rgba(0,0,0,0.2);
                z-index: 9999;
                max-width: 300px;
                transition: all 0.3s ease;
            `;
            notification.textContent = message;
            document.body.appendChild(notification);

            setTimeout(() => {
                notification.style.opacity = '0';
                setTimeout(() => notification.remove(), 300);
            }, duration);
        },

        // Esperar un tiempo
        sleep(ms) {
            return new Promise(resolve => setTimeout(resolve, ms));
        }
    };

    // Funciones de API
    const api = {
        // Obtener grupos de tareas existentes
        async getAssignmentGroups(courseId) {
            try {
                const response = await fetch(`/api/v1/courses/${courseId}/assignment_groups`, {
                    headers: { 'Accept': 'application/json' },
                    credentials: 'include'
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.message || 'Error al obtener grupos');
                }

                return await response.json();
            } catch (error) {
                console.error('Error en getAssignmentGroups:', error);
                throw error;
            }
        },

        // Crear nuevo grupo de tareas
        async createAssignmentGroup(courseId, groupName) {
            try {
                const response = await fetch(`/api/v1/courses/${courseId}/assignment_groups`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRF-Token': state.csrfToken
                    },
                    body: JSON.stringify({
                        name: groupName,
                        group_weight: 0
                    }),
                    credentials: 'include'
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.message || 'Error al crear grupo');
                }

                return await response.json();
            } catch (error) {
                console.error('Error en createAssignmentGroup:', error);
                throw error;
            }
        },

        // Crear nueva tarea
        async createAssignment(courseId, assignmentData) {
            try {
                const response = await fetch(`/api/v1/courses/${courseId}/assignments`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRF-Token': state.csrfToken
                    },
                    body: JSON.stringify({
                        assignment: assignmentData
                    }),
                    credentials: 'include'
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.message || 'Error al crear tarea');
                }

                return await response.json();
            } catch (error) {
                console.error('Error en createAssignment:', error);
                throw error;
            }
        },

        // Eliminar grupo de tareas
        async deleteAssignmentGroup(courseId, groupId) {
            try {
                const response = await fetch(`/api/v1/courses/${courseId}/assignment_groups/${groupId}`, {
                    method: 'DELETE',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRF-Token': state.csrfToken
                    },
                    credentials: 'include'
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.message || 'Error al eliminar grupo');
                }

                return await response.json();
            } catch (error) {
                console.error('Error en deleteAssignmentGroup:', error);
                throw error;
            }
        }
    };

    // Funciones principales
    const main = {
        // Inicializar el script
        async init() {
            state.courseId = utils.getCourseId();
            state.csrfToken = utils.getCSRFToken();

            if (!state.courseId) {
                utils.showNotification('No se pudo obtener el ID del curso. Asegúrate de estar en la página de tareas del curso.', 'error');
                console.error('No se pudo obtener courseId. URL actual:', window.location.href);
                return;
            }

            if (!state.csrfToken) {
                utils.showNotification('Error de autenticación. Por favor recarga la página.', 'error');
                console.error('No se pudo obtener CSRF token');
                return;
            }

            this.addButton();
        },

        // Añadir botón a la interfaz
        addButton() {
            const container = document.querySelector('div[role="main"] > div > div');
            if (!container || document.getElementById('btnCrearEstructura')) return;

            const btn = document.createElement('button');
            btn.id = 'btnCrearEstructura';
            btn.textContent = '📘 Crear estructura de tareas';
            btn.style.backgroundColor = '#007bff';
            btn.style.color = 'white';
            btn.style.border = 'none';
            btn.style.padding = '8px 16px';
            btn.style.marginLeft = '8px';
            btn.style.borderRadius = '4px';
            btn.style.cursor = 'pointer';
            btn.onclick = () => this.runProcess();

            const btnDelete = document.createElement('button');
            btnDelete.id = 'btnEliminarEstructura';
            btnDelete.textContent = '🗑️ Eliminar estructura';
            btnDelete.style.backgroundColor = '#dc3545';
            btnDelete.style.color = 'white';
            btnDelete.style.border = 'none';
            btnDelete.style.padding = '8px 16px';
            btnDelete.style.marginLeft = '8px';
            btnDelete.style.borderRadius = '4px';
            btnDelete.style.cursor = 'pointer';
            btnDelete.onclick = () => this.deleteAllGroupsExceptFirst();

            const actionsBar = document.querySelector('[role="main"] .header-bar-right') || container;
            actionsBar.appendChild(btn);
            actionsBar.appendChild(btnDelete);
        },

        // Ejecutar el proceso completo de creación
        async runProcess() {
            if (state.isRunning) return;
            state.isRunning = true;

            const btn = document.getElementById('btnCrearEstructura');
            if (btn) btn.disabled = true;

            utils.showNotification('Iniciando creación de estructura...', 'info');

            try {
                for (const [index, item] of config.estructura.entries()) {
                    const progress = `(${index + 1}/${config.estructura.length})`;
                    utils.showNotification(`${progress} Procesando: ${item.tarea}`, 'info');

                    // 1. Obtener o crear grupo
                    let groupId = await this.getOrCreateGroup(item.grupo);

                    if (groupId) {
                        // 2. Crear tarea
                        await this.createAssignmentInGroup(groupId, item.tarea);

                        // Esperar entre solicitudes
                        await utils.sleep(config.delayBetweenRequests);
                    }
                }

                utils.showNotification('¡Estructura creada con éxito!', 'success', 3000);
            } catch (error) {
                utils.showNotification(`Error: ${error.message}`, 'error');
                console.error('Error en el proceso:', error);
            } finally {
                state.isRunning = false;
                if (btn) btn.disabled = false;

                // Recargar después de 3 segundos
                setTimeout(() => location.reload(), 3000);
            }
        },

        // Obtener o crear grupo de tareas
        async getOrCreateGroup(groupName) {
            try {
                // 1. Verificar si el grupo ya existe
                const groups = await api.getAssignmentGroups(state.courseId);
                const existingGroup = groups.find(g => g.name === groupName);

                if (existingGroup) {
                    console.log(`Grupo existente: ${groupName} (ID: ${existingGroup.id})`);
                    return existingGroup.id;
                }

                // 2. Crear nuevo grupo
                utils.showNotification(`Creando grupo: ${groupName}`, 'info');
                const newGroup = await api.createAssignmentGroup(state.courseId, groupName);

                console.log(`Nuevo grupo creado: ${groupName} (ID: ${newGroup.id})`);
                return newGroup.id;
            } catch (error) {
                utils.showNotification(`Error con grupo ${groupName}: ${error.message}`, 'error');
                throw error;
            }
        },

        // Crear tarea en un grupo
        async createAssignmentInGroup(groupId, assignmentName) {
            try {
                const assignmentData = {
                    name: assignmentName,
                    assignment_group_id: groupId,
                    points_possible: config.puntosPorTarea,
                    submission_types: ['online_text_entry', 'online_url', 'online_upload'],
                    published: false,
                    grading_type: 'points',
                    omit_from_final_grade: false
                };

                utils.showNotification(`Creando tarea: ${assignmentName}`, 'info');
                const assignment = await api.createAssignment(state.courseId, assignmentData);

                console.log(`Tarea creada: ${assignmentName} (ID: ${assignment.id})`);
                return assignment;
            } catch (error) {
                utils.showNotification(`Error con tarea ${assignmentName}: ${error.message}`, 'error');
                throw error;
            }
        },

        // Función para eliminar todos los grupos excepto el protegido
        async deleteAllGroupsExceptFirst() {
            if (state.isRunning) return;
            state.isRunning = true;

            const btnDelete = document.getElementById('btnEliminarEstructura');
            if (btnDelete) btnDelete.disabled = true;

            if (!confirm(`¿Estás seguro de que deseas eliminar todos los grupos de evaluación excepto "${config.grupoProtegido}"? Esta acción no se puede deshacer.`)) {
                state.isRunning = false;
                if (btnDelete) btnDelete.disabled = false;
                return;
            }

            utils.showNotification('Iniciando eliminación de grupos...', 'info');

            try {
                // 1. Obtener todos los grupos
                const groups = await api.getAssignmentGroups(state.courseId);
                
                // 2. Filtrar los grupos que no son el protegido y están en la estructura
                const groupsToDelete = groups.filter(group => 
                    group.name !== config.grupoProtegido && 
                    config.estructura.some(item => item.grupo === group.name)
                );

                if (groupsToDelete.length === 0) {
                    utils.showNotification('No se encontraron grupos para eliminar', 'info');
                    return;
                }

                // 3. Eliminar cada grupo
                for (const [index, group] of groupsToDelete.entries()) {
                    const progress = `(${index + 1}/${groupsToDelete.length})`;
                    utils.showNotification(`${progress} Eliminando grupo: ${group.name}`, 'info');
                    
                    await api.deleteAssignmentGroup(state.courseId, group.id);
                    await utils.sleep(config.delayBetweenRequests);
                }

                utils.showNotification('¡Grupos eliminados con éxito!', 'success', 3000);
            } catch (error) {
                utils.showNotification(`Error: ${error.message}`, 'error');
                console.error('Error al eliminar grupos:', error);
            } finally {
                state.isRunning = false;
                if (btnDelete) btnDelete.disabled = false;

                // Recargar después de 3 segundos
                setTimeout(() => location.reload(), 3000);
            }
        }
    };

    // Inicialización
    if (document.readyState === 'complete') {
        main.init();
        window.main = main;
    } else {
        window.addEventListener('load', () => main.init());
    }
})();