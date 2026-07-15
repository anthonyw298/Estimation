'use client';

import { AlertTriangle, AlertCircle, CheckCircle } from 'lucide-react';
import { cn } from '@/lib/utils';

interface ValidationResult {
  valid: boolean;
  message: string;
  glass_after?: number;
  door_area?: number;
  is_warning?: boolean;
}

interface DoorValidatorProps {
  result: ValidationResult | null;
  className?: string;
}

export function DoorValidator({ result, className }: DoorValidatorProps) {
  if (!result) return null;

  const isCaution = !result.valid && !result.is_warning;
  const isWarning = result.is_warning;
  const isSuccess = result.valid && !result.is_warning;

  return (
    <div
      className={cn(
        'mt-3 p-3 rounded-lg flex gap-3 items-start',
        isCaution && 'bg-red-50 border border-red-200',
        isWarning && 'bg-amber-50 border border-amber-200',
        isSuccess && 'bg-green-50 border border-green-200',
        className
      )}
    >
      {isCaution && <AlertTriangle className="w-5 h-5 text-red-600 flex-shrink-0 mt-0.5" />}
      {isWarning && <AlertCircle className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" />}
      {isSuccess && <CheckCircle className="w-5 h-5 text-green-600 flex-shrink-0 mt-0.5" />}

      <div className="flex-1">
        <p
          className={cn(
            'text-sm font-medium',
            isCaution && 'text-red-800',
            isWarning && 'text-amber-800',
            isSuccess && 'text-green-800'
          )}
        >
          {result.message}
        </p>
        {result.door_area && result.glass_after !== undefined && (
          <p className="text-xs mt-1 opacity-75">
            Door: {result.door_area.toFixed(2)} sqft | Remaining: {result.glass_after.toFixed(2)} sqft
          </p>
        )}
      </div>
    </div>
  );
}
