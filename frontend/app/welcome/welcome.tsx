'use client';

import { useState, useEffect, useCallback } from 'react';
import ReactFlow, {
	Controls,
	Background,
	useNodesState,
	useEdgesState,
	Handle,
	Position,
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
import { EtapaNombre } from '~/config/enum';
import ExtraccionNode from '~/components/flow/extraccion';
import { formatTime } from '~/lib/utils';
import LimpiezaNode from '~/components/flow/limpieza';
import GuardadoNode from '~/components/flow/guardado';

interface FileInfo {
	archivo_id: string;
	nombre: string;
	filas: number;
	fecha_creacion?: string | null;
	url?: string;
}

interface ProcessingStep {
	id: number;
	nombre: string;
	orden: number;
	filas: number;
	status: string;
	duracion_ms: number | null;
	detalles: string | null;
	mensaje_error: string | null;
	registros_procesados: number;
	inicio: string;
	fin: string;
}

interface FlowData {
	archivo?: {
		url?: string;
	};
	pasos: ProcessingStep[];
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
				const listadoRaw = (data.archivos || []) as any[];
				const listado: FileInfo[] = listadoRaw.map((a) => ({
					archivo_id: a.archivo_id,
					nombre: a.nombre,
					filas: Number(a.filas) || 0,
					url: a.url || a.webUrl || null,
					fecha_creacion: a.fecha_creacion || null,
				}));
				setArchivos(listado);

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
					const dataRaw = await response.json();
					const data: FlowData = {
						archivo: dataRaw.archivo || null,
						pasos: (dataRaw.pasos || []).map((p: any) => {
							const rawStatus = (p.estado || p.status || '').toString().toLowerCase();
							let statusNorm = 'pendiente';
							if (rawStatus.includes('proces')) statusNorm = 'procesando';
							else if (rawStatus.includes('complet')) statusNorm = 'completado';
							else if (rawStatus.includes('fall') || rawStatus.includes('error')) statusNorm = 'error';

							return {
								id: p.id,
								nombre: p.nombre,
								orden: p.orden,
								filas: Number(p.filas) || 0,
								status: statusNorm,
								duracion_ms: p.duracion_ms ?? null,
								detalles: p.detalles ?? null,
								mensaje_error: p.mensaje_error ?? null,
								registros_procesados: p.registros_procesados ?? 0,
								inicio: p.inicio ?? null,
								fin: p.fin ?? null,
							};
						}),
					};
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
									filas={archivoSeleccionado?.filas ?? 0}
								/>
							),
						},
						position: { x: 0, y: 160 },
						sourcePosition: Position.Right,
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
								label: (
									paso.nombre == EtapaNombre.EXTRACCION ? (
										<ExtraccionNode 
											nombre={paso.nombre}
											filas={paso.filas}
											tiempo={formatTime(paso.inicio ?? '', paso.fin ?? '')}
										/>
									) : (paso.nombre === EtapaNombre.LIMPIEZA ? (
											<LimpiezaNode
												nombre={paso.nombre}
												filas={paso.filas}
												tiempo={formatTime(paso.inicio ?? '', paso.fin ?? '')}
											/>
										) : ( paso.nombre === EtapaNombre.GUARDADO ? (
												<GuardadoNode 
													nombre={paso.nombre}
													filas={paso.filas}
													tiempo={formatTime(paso.inicio ?? '', paso.fin ?? '')}
											/>
										)
											: (
												<>
													<Handle type="target" position={Position.Left} />
													<span>{paso.nombre.charAt(0).toUpperCase() + paso.nombre.slice(1)}</span>
													<Handle type="source" position={Position.Right} />
												</>
											)
										)
									)	
								),
							},
							position: { x: 240 * (index + 1), y: 160 },
							style: {
								background: '#ffffff',
								border: '1px solid #cbd5e1',
								borderRadius: 14,
								padding: 0,
								width: 200,
								boxShadow: '0 10px 25px rgba(15, 23, 42, 0.08)',
							},
						});

						if (index === 0) {
							newEdges.push({
								id: `edge-inicio-${paso.id}`,
								source: 'inicio',
								target: pasoNodeId,
								animated: paso.status === 'procesando',
								type: 'straight',
							});
						} else {
							const prevPaso = data.pasos[index - 1];
							newEdges.push({
								id: `edge-${prevPaso.id}-${paso.id}`,
								source: `paso-${prevPaso.id}`,
								target: pasoNodeId,
								animated: paso.status === 'procesando',
								type: 'straight',
							});
						}
					});
					
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
			<header className="fixed top-4 left-4 z-10 mx-auto flex items-center gap-4">
				<Select value={selectedFileId} onValueChange={setSelectedFileId}>
					<SelectTrigger className="min-w-90">
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
