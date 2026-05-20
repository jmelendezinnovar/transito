import { ArrowDownToLine, BrushCleaning, FileSpreadsheet } from "lucide-react";
import { EtapaNombre, getEtapaNombreDisplay } from "~/config/enum";

interface ExcelNodeProps {
    nombre: string;
    filas: number;
    tiempo: string;
}

export default function LimpiezaNode({ nombre, filas, tiempo }: ExcelNodeProps) {
    return (
        <div className="h-full w-full overflow-hidden rounded-xl bg-white">
            <div className="border-b border-slate-200 px-3 py-2 text-left text-xs font-semibold text-slate-800">
                {getEtapaNombreDisplay(nombre as EtapaNombre)}
            </div>
            <div className="grid grid-cols-[auto_1fr] items-center gap-2 px-3 py-2 text-start">
                <div className="flex h-12 w-12 items-center justify-center rounded-md bg-blue-300 text-white">
                    <BrushCleaning size={26} />
                </div>
                <div className="grid grid-cols-1 gap-1 h-auto">
                    <div className="truncate text-xs font-semibold text-slate-800">
                        {filas} filas
                    </div>
                    <div className="">
                        <span className="text-xs text-slate-700">
                            {tiempo}
                        </span>
                    </div>
                </div>
            </div>
        </div>
    );
}