import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export function formatTime(inicio: string, fin: string) {
	const inicioTime = new Date(inicio);
	const finTime = new Date(fin);
	let diffInSeconds = Math.floor((finTime.getTime() - inicioTime.getTime()) / 1000);
	if (isNaN(diffInSeconds) || diffInSeconds < 0) diffInSeconds = 0;

	const plural = (n: number, singular: string, pluralForm: string) =>
		`${n} ${n === 1 ? singular : pluralForm}`;

	if (diffInSeconds < 60) {
		return plural(diffInSeconds, 'segundo', 'segundos');
	}

	if (diffInSeconds < 3600) {
		const minutes = Math.floor(diffInSeconds / 60);
		const seconds = diffInSeconds % 60;
		if (seconds === 0) return plural(minutes, 'minuto', 'minutos');
		return `${plural(minutes, 'minuto', 'minutos')} ${plural(seconds, 'segundo', 'segundos')}`;
	}

	if (diffInSeconds < 86400) {
		const hours = Math.floor(diffInSeconds / 3600);
		const minutes = Math.floor((diffInSeconds % 3600) / 60);
		if (minutes === 0) return plural(hours, 'hora', 'horas');
		return `${plural(hours, 'hora', 'horas')} ${plural(minutes, 'minuto', 'minutos')}`;
	}

	const days = Math.floor(diffInSeconds / 86400);
	const hours = Math.floor((diffInSeconds % 86400) / 3600);
	if (hours === 0) return plural(days, 'día', 'días');
	return `${plural(days, 'día', 'días')} ${plural(hours, 'hora', 'horas')}`;
}