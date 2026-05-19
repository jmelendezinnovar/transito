'use client';

import { useState, useEffect, useCallback } from 'react';
import ReactFlow, {
	Controls,
	Background,
	useNodesState,
	useEdgesState,
} from 'reactflow';
import type { Node, Edge } from 'reactflow';
import 'reactflow/dist/style.css';
import { API_ROUTES } from '../config/api';
import {
	Select,
	SelectContent,
	SelectItem,
	SelectTrigger,
	SelectValue,
} from '~/components/ui/select';
import ExcelNode from '~/components/flow/excel';

interface FileInfo {
	archivo_id: string;
	nombre: string;
	url?: string;
}

interface ProcessingStep {
	id: number;
	nombre: string;
	orden: number;
	status: string;
	duracion_ms: number | null;
	detalles: string | null;
	mensaje_error: string | null;
	registros_procesados: number;
	inicio: string | null;
	fin: string | null;
}

interface FlowData {
	archivo?: {
		url?: string;
	};
	pasos: ProcessingStep[];
}

function nodeStyleByStatus(status: string): React.CSSProperties {
	if (status === 'completado') {
		return {
			background: '#ecfdf3',
			border: '2px solid #16a34a',
			color: '#14532d',
			borderRadius: 12,
			padding: 8,
			fontWeight: 600,
			minWidth: 180,
			textAlign: 'center',
		};
	}

	if (status === 'error') {
		return {
			background: '#fef2f2',
			border: '2px solid #dc2626',
			color: '#7f1d1d',
			borderRadius: 12,
			padding: 8,
			fontWeight: 600,
			minWidth: 180,
			textAlign: 'center',
		};
	}

	return {
		background: '#eff6ff',
		border: '2px solid #2563eb',
		color: '#1e3a8a',
		borderRadius: 12,
		padding: 8,
		fontWeight: 600,
		minWidth: 180,
		textAlign: 'center',
	};
}

export function Welcome() {
	const [archivos, setArchivos] = useState<FileInfo[]>([]);
	const [selectedFileId, setSelectedFileId] = useState<string>('');
	const [loading, setLoading] = useState(false);
	const [error, setError] = useState<string>('');
	const [nodes, setNodes, onNodesChange] = useNodesState([]);
	const [edges, setEdges, onEdgesChange] = useEdgesState([]);
	const archivoSeleccionado = archivos.find(
		(archivo) => archivo.archivo_id === selectedFileId
	);

	const cargarArchivos = useCallback(async () => {
		setLoading(true);
		setError('');
		try {
			const response = await fetch(API_ROUTES.archivos, {
				headers: {
					'Content-Type': 'application/json',
				},
			});

			if (response.ok) {
				const data = await response.json();
				const listado: FileInfo[] = data.archivos || [];
				setArchivos(listado);

				console.log(listado);

				if (!selectedFileId && listado.length > 0) {
					setSelectedFileId(listado[0].archivo_id);
				}
			} else {
				setError('No se pudo cargar la lista de archivos.');
			}
		} catch (error) {
			console.error('Error cargando archivos:', error);
			setError('Error de conexión al cargar archivos.');
		} finally {
			setLoading(false);
		}
	}, [selectedFileId]);

	const cargarFlujo = useCallback(
		async (archivoId: string) => {
			if (!archivoId) {
				setNodes([]);
				setEdges([]);
				return;
			}

			setLoading(true);
			setError('');
			try {
				const response = await fetch(API_ROUTES.archivoFlujo(archivoId), {
					headers: {
						'Content-Type': 'application/json',
					},
				});

				if (response.ok) {
					const data: FlowData = await response.json();
					const newNodes: Node[] = [];
					const newEdges: Edge[] = [];
					const excelUrl = data.archivo?.url ?? archivoSeleccionado?.url;

					newNodes.push({
						id: 'inicio',
						data: {
							label: (
								<ExcelNode
									nombre={archivoSeleccionado?.nombre ?? 'Archivo cargado'}
									url={excelUrl}
									filas={1000}
								/>
							),
						},
						position: { x: 0, y: 160 },
						style: {
							background: '#ffffff',
							border: '1px solid #cbd5e1',
							borderRadius: 14,
							padding: 0,
							width: 200,
							boxShadow: '0 10px 25px rgba(15, 23, 42, 0.08)',
						},
					});

					data.pasos.forEach((paso, index) => {
						const pasoNodeId = `paso-${paso.id}`;

						newNodes.push({
							id: pasoNodeId,
							data: {
								label: paso.nombre.charAt(0).toUpperCase() + paso.nombre.slice(1),
							},
							position: { x: 240 * (index + 1), y: 160 },
							style: nodeStyleByStatus(paso.status),
						});

						if (index === 0) {
							newEdges.push({
								id: `edge-inicio-${paso.id}`,
								source: 'inicio',
								target: pasoNodeId,
								animated: paso.status === 'procesando',
							});
						} else {
							const prevPaso = data.pasos[index - 1];
							newEdges.push({
								id: `edge-${prevPaso.id}-${paso.id}`,
								source: `paso-${prevPaso.id}`,
								target: pasoNodeId,
								animated: paso.status === 'procesando',
							});
						}
					});

					const finX = 240 * (data.pasos.length + 1);
					const finStatus = data.pasos.some((paso) => paso.status === 'error')
						? 'error'
						: 'completado';

					newNodes.push({
						id: 'fin',
						data: { label: 'Fin' },
						position: { x: finX, y: 160 },
						style: nodeStyleByStatus(finStatus),
					});

					if (data.pasos.length > 0) {
						const lastPaso = data.pasos[data.pasos.length - 1];
						newEdges.push({
							id: `edge-${lastPaso.id}-fin`,
							source: `paso-${lastPaso.id}`,
							target: 'fin',
							animated: lastPaso.status === 'procesando',
						});
					} else {
						newEdges.push({
							id: 'edge-inicio-fin',
							source: 'inicio',
							target: 'fin',
						});
					}

					setNodes(newNodes);
					setEdges(newEdges);
				} else {
					setError('No se pudo cargar el flujo del archivo seleccionado.');
				}
			} catch (error) {
				console.error('Error cargando flujo:', error);
				setError('Error de conexión al cargar el flujo.');
			} finally {
				setLoading(false);
			}
		},
		[setNodes, setEdges, archivoSeleccionado?.nombre, archivoSeleccionado?.url]
	);

	useEffect(() => {
		cargarArchivos();
	}, [cargarArchivos]);

	useEffect(() => {
		if (selectedFileId) {
			cargarFlujo(selectedFileId);
		}
	}, [selectedFileId, cargarFlujo]);

	return (
		<div className="w-screen h-screen grid grid-cols-1 grid-rows-1 overflow-hidden relative bg-white">
			<header className="fixed top-[16px] left-[16px] z-10 mx-auto flex items-center gap-4">
				<Select value={selectedFileId} onValueChange={setSelectedFileId}>
					<SelectTrigger className="min-w-[360px]">
						<SelectValue placeholder="Selecciona un archivo" />
					</SelectTrigger>
					<SelectContent>
						{archivos.length === 0 ? (
							<SelectItem value="__sin_archivos" disabled>
								No hay archivos
							</SelectItem>
						) : (
							archivos.map((archivo) => (
								<SelectItem key={archivo.archivo_id} value={archivo.archivo_id}>
									{archivo.nombre}
								</SelectItem>
							))
						)}
					</SelectContent>
				</Select>

				{loading && <span className="text-sm text-slate-500">Cargando...</span>}
				{error && <span className="text-sm text-red-600">{error}</span>}
			</header>

			<main className="w-full h-full">
				<ReactFlow
					nodes={nodes}
					edges={edges}
					onNodesChange={onNodesChange}
					onEdgesChange={onEdgesChange}
					fitView
					fitViewOptions={{ padding: 0.25 }}
				>
					<Background color="#e2e8f0" gap={24} />
					<Controls position="bottom-right" />
				</ReactFlow>
			</main>
		</div>
	);
}
