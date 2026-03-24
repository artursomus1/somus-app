import React, { useState, useMemo } from 'react';
import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  flexRender,
  type ColumnDef,
  type SortingState,
  type RowSelectionState,
} from '@tanstack/react-table';
import {
  Search,
  ChevronUp,
  ChevronDown,
  ChevronsUpDown,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  Download,
} from 'lucide-react';
import { cn } from '@/utils/cn';

export interface DataTableProps<T> {
  data: T[];
  columns: ColumnDef<T, any>[];
  searchable?: boolean;
  paginated?: boolean;
  pageSize?: number;
  selectable?: boolean;
  onRowClick?: (row: T) => void;
  exportFilename?: string;
  className?: string;
  emptyMessage?: string;
}

function exportToCSV<T>(data: T[], columns: ColumnDef<T, any>[], filename: string) {
  const headers = columns
    .map((col: any) => col.header?.toString() ?? col.id ?? '')
    .filter(Boolean);

  const rows = data.map((row: any) =>
    columns.map((col: any) => {
      const key = col.accessorKey ?? col.id;
      const val = key ? row[key] : '';
      // Format numbers to Brazilian style for CSV
      if (typeof val === 'number') {
        return val.toLocaleString('pt-BR');
      }
      // Escape quotes in strings
      if (typeof val === 'string' && (val.includes(';') || val.includes('"'))) {
        return `"${val.replace(/"/g, '""')}"`;
      }
      return val ?? '';
    })
  );

  const csvContent = [headers.join(';'), ...rows.map((r) => r.join(';'))].join('\n');

  // BOM for UTF-8
  const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

export function DataTable<T>({
  data,
  columns,
  searchable = true,
  paginated = true,
  pageSize = 15,
  selectable = false,
  onRowClick,
  exportFilename,
  className,
  emptyMessage = 'Nenhum registro encontrado.',
}: DataTableProps<T>) {
  const [sorting, setSorting] = useState<SortingState>([]);
  const [globalFilter, setGlobalFilter] = useState('');
  const [rowSelection, setRowSelection] = useState<RowSelectionState>({});

  // Prepend checkbox column if selectable
  const finalColumns = useMemo(() => {
    if (!selectable) return columns;

    const checkboxCol: ColumnDef<T, any> = {
      id: 'select',
      header: ({ table }) => (
        <input
          type="checkbox"
          checked={table.getIsAllPageRowsSelected()}
          onChange={table.getToggleAllPageRowsSelectedHandler()}
          className="rounded border-somus-gray-300 text-somus-green focus:ring-somus-green/40"
        />
      ),
      cell: ({ row }) => (
        <input
          type="checkbox"
          checked={row.getIsSelected()}
          onChange={row.getToggleSelectedHandler()}
          className="rounded border-somus-gray-300 text-somus-green focus:ring-somus-green/40"
          onClick={(e) => e.stopPropagation()}
        />
      ),
      size: 40,
      enableSorting: false,
    };

    return [checkboxCol, ...columns];
  }, [columns, selectable]);

  const table = useReactTable({
    data,
    columns: finalColumns,
    state: {
      sorting,
      globalFilter,
      rowSelection,
    },
    onSortingChange: setSorting,
    onGlobalFilterChange: setGlobalFilter,
    onRowSelectionChange: setRowSelection,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getPaginationRowModel: paginated ? getPaginationRowModel() : undefined,
    enableRowSelection: selectable,
    initialState: {
      pagination: { pageSize },
    },
  });

  const { pageIndex, pageSize: currentPageSize } = table.getState().pagination;
  const totalRows = table.getFilteredRowModel().rows.length;

  return (
    <div className={cn('flex flex-col gap-3', className)}>
      {/* ── Toolbar ── */}
      {(searchable || exportFilename) && (
        <div className="flex items-center justify-between gap-3">
          {searchable && (
            <div className="relative max-w-xs w-full">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-somus-gray-400" />
              <input
                type="text"
                value={globalFilter}
                onChange={(e) => setGlobalFilter(e.target.value)}
                placeholder="Pesquisar..."
                className="w-full rounded-lg border border-somus-gray-300 bg-white py-2 pl-9 pr-3 text-sm text-somus-gray-900 placeholder:text-somus-gray-400 focus:outline-none focus:ring-2 focus:ring-somus-green/40 focus:border-somus-green transition-colors"
              />
            </div>
          )}

          {exportFilename && (
            <button
              onClick={() => exportToCSV(data, columns, exportFilename)}
              className="inline-flex items-center gap-2 px-3 py-2 text-sm font-medium text-somus-gray-600 bg-white border border-somus-gray-300 rounded-lg hover:bg-somus-gray-50 transition-colors shadow-sm"
            >
              <Download className="h-4 w-4" />
              Exportar CSV
            </button>
          )}
        </div>
      )}

      {/* ── Table ── */}
      <div className="overflow-x-auto rounded-lg border border-somus-gray-200 bg-white">
        <table className="w-full text-sm">
          <thead>
            {table.getHeaderGroups().map((headerGroup) => (
              <tr key={headerGroup.id} className="border-b border-somus-gray-200 bg-somus-gray-50">
                {headerGroup.headers.map((header) => (
                  <th
                    key={header.id}
                    onClick={header.column.getCanSort() ? header.column.getToggleSortingHandler() : undefined}
                    className={cn(
                      'px-4 py-3 text-left text-xs font-semibold text-somus-gray-600 uppercase tracking-wider whitespace-nowrap',
                      header.column.getCanSort() && 'cursor-pointer select-none hover:text-somus-gray-900'
                    )}
                    style={{ width: header.getSize() !== 150 ? header.getSize() : undefined }}
                  >
                    <div className="flex items-center gap-1">
                      {header.isPlaceholder
                        ? null
                        : flexRender(header.column.columnDef.header, header.getContext())}
                      {header.column.getCanSort() && (
                        <span className="ml-1">
                          {header.column.getIsSorted() === 'asc' ? (
                            <ChevronUp className="h-3.5 w-3.5" />
                          ) : header.column.getIsSorted() === 'desc' ? (
                            <ChevronDown className="h-3.5 w-3.5" />
                          ) : (
                            <ChevronsUpDown className="h-3.5 w-3.5 text-somus-gray-300" />
                          )}
                        </span>
                      )}
                    </div>
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody className="divide-y divide-somus-gray-100">
            {table.getRowModel().rows.length === 0 ? (
              <tr>
                <td
                  colSpan={finalColumns.length}
                  className="px-4 py-12 text-center text-somus-gray-400"
                >
                  {emptyMessage}
                </td>
              </tr>
            ) : (
              table.getRowModel().rows.map((row) => (
                <tr
                  key={row.id}
                  onClick={() => onRowClick?.(row.original)}
                  className={cn(
                    'transition-colors',
                    onRowClick && 'cursor-pointer',
                    row.getIsSelected()
                      ? 'bg-somus-green-bg'
                      : 'hover:bg-somus-gray-50'
                  )}
                >
                  {row.getVisibleCells().map((cell) => (
                    <td
                      key={cell.id}
                      className="px-4 py-3 text-somus-gray-700 whitespace-nowrap"
                    >
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* ── Pagination ── */}
      {paginated && totalRows > currentPageSize && (
        <div className="flex items-center justify-between text-sm text-somus-gray-500">
          <span>
            Mostrando {pageIndex * currentPageSize + 1} a{' '}
            {Math.min((pageIndex + 1) * currentPageSize, totalRows)} de {totalRows}{' '}
            registros
          </span>

          <div className="flex items-center gap-1">
            <button
              onClick={() => table.setPageIndex(0)}
              disabled={!table.getCanPreviousPage()}
              className="p-1.5 rounded hover:bg-somus-gray-100 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
            >
              <ChevronsLeft className="h-4 w-4" />
            </button>
            <button
              onClick={() => table.previousPage()}
              disabled={!table.getCanPreviousPage()}
              className="p-1.5 rounded hover:bg-somus-gray-100 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>

            <span className="px-3 py-1 text-sm font-medium">
              {pageIndex + 1} / {table.getPageCount()}
            </span>

            <button
              onClick={() => table.nextPage()}
              disabled={!table.getCanNextPage()}
              className="p-1.5 rounded hover:bg-somus-gray-100 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
            >
              <ChevronRight className="h-4 w-4" />
            </button>
            <button
              onClick={() => table.setPageIndex(table.getPageCount() - 1)}
              disabled={!table.getCanNextPage()}
              className="p-1.5 rounded hover:bg-somus-gray-100 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
            >
              <ChevronsRight className="h-4 w-4" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default DataTable;
