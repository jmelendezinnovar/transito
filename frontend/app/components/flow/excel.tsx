import { FileSpreadsheet } from "lucide-react";

interface ExcelNodeProps {
    nombre: string;
    url?: string;
    filas: number;
}

export default function ExcelNode({nombre, url, filas}: ExcelNodeProps) {
    return (
        <div className="h-full w-full overflow-hidden rounded-xl bg-white">
            <div className="border-b border-slate-200 px-3 py-2 text-left text-xs font-semibold text-slate-800">
                {nombre}
            </div>
            <div className="grid grid-cols-[auto_1fr] items-center gap-2 px-3 py-2 text-start">
                <div className="flex h-12 w-12 items-center justify-center rounded-md bg-emerald-600 text-white">
                    <FileSpreadsheet size={26} />
                </div>
                <div className="grid grid-cols-1 gap-1 h-auto">
                    <div className="ml-1 truncate text-xs font-semibold text-slate-800">
                        {filas} filas
                    </div>
                    <div className="">
                        {url ? (
                            <a
                                href={url}
                                target="_blank"
                                rel="noreferrer"
                                className="rounded-md bg-emerald-100 px-2 py-1 text-xs font-semibold text-emerald-700 hover:bg-emerald-200"
                            >
                                Abrir Excel
                            </a>
                        ) : (
                            <span className="text-xs text-slate-500">Sin URL</span>
                        )}
                    </div>
                </div>
            </div>
        </div>
    );
}