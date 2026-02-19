'use client';

import { Plus, Trash2, DoorOpen } from 'lucide-react';
import type { DoorConfig } from '@/types';

interface DoorManagerProps {
  doors: DoorConfig[];
  onChange: (doors: DoorConfig[]) => void;
  finish: string;
}

const DOOR_SIZES = [
  "3' X 7'",
  "3' X 8'",
  "3' X 9'",
  "6' X 7'",
  "6' X 8'",
  "6' X 9'",
];

const STILE_OPTIONS = ['Narrow', 'Medium', 'Wide'];

const HARDWARE_OPTIONS = [
  'Concealed Closer',
  'Exit Devices',
  'Continuous Hinges',
  'Latch Lock w/ Lever Handle',
  'Lever Handle',
  'Electric Strike',
  'Extended Ladder Pull (B2B)',
  'Extended Ladder Pull (Single)',
];

export default function DoorManager({ doors, onChange, finish }: DoorManagerProps) {
  function addDoor() {
    const newDoor: DoorConfig = {
      size: DOOR_SIZES[0],
      count: 1,
      stile: 'Narrow',
      hardware: {},
    };
    onChange([...doors, newDoor]);
  }

  function updateDoor(index: number, updates: Partial<DoorConfig>) {
    const updated = doors.map((door, i) =>
      i === index ? { ...door, ...updates } : door
    );
    onChange(updated);
  }

  function removeDoor(index: number) {
    onChange(doors.filter((_, i) => i !== index));
  }

  function toggleHardware(doorIndex: number, hardwareName: string) {
    const door = doors[doorIndex];
    const newHardware = { ...door.hardware };
    newHardware[hardwareName] = !newHardware[hardwareName];
    updateDoor(doorIndex, { hardware: newHardware });
  }

  return (
    <div className="space-y-4">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <DoorOpen className="w-5 h-5 text-[#2563eb]" />
          <h3 className="text-sm font-semibold text-white tracking-tight">
            Doors
          </h3>
          <span className="text-xs text-[#71717a]">
            ({doors.length} door{doors.length !== 1 ? 's' : ''})
          </span>
        </div>
        <button
          onClick={addDoor}
          className="flex items-center gap-1.5 px-3 py-1.5 bg-[#2563eb] hover:bg-[#3b82f6] text-white text-xs font-medium rounded-md transition-colors duration-150"
        >
          <Plus className="w-3.5 h-3.5" />
          Add Door
        </button>
      </div>

      {/* Door cards */}
      {doors.length === 0 && (
        <div className="border border-dashed border-[#27272a] rounded-lg p-6 text-center">
          <DoorOpen className="w-8 h-8 text-[#3f3f46] mx-auto mb-2" />
          <p className="text-sm text-[#71717a]">
            No doors added yet. Click &quot;Add Door&quot; to get started.
          </p>
        </div>
      )}

      {doors.map((door, index) => (
        <div
          key={index}
          className="bg-[#18181b] border border-[#27272a] rounded-lg p-4 space-y-4"
        >
          {/* Door card header */}
          <div className="flex items-center justify-between">
            <span className="text-sm font-medium text-white">
              Door {index + 1}
            </span>
            <button
              onClick={() => removeDoor(index)}
              className="p-1.5 text-[#71717a] hover:text-red-400 hover:bg-red-400/10 rounded-md transition-colors duration-150"
              title="Remove door"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          </div>

          {/* Size, Count, Stile row */}
          <div className="grid grid-cols-3 gap-3">
            {/* Size */}
            <div>
              <label className="block text-xs text-[#a1a1aa] mb-1 font-medium">
                Size
              </label>
              <select
                value={door.size}
                onChange={(e) => updateDoor(index, { size: e.target.value })}
                className="w-full px-2.5 py-1.5 bg-[#09090b] border border-[#27272a] rounded-md text-sm text-white focus:outline-none focus:border-[#2563eb] transition-colors"
              >
                {DOOR_SIZES.map((size) => (
                  <option key={size} value={size}>
                    {size}
                  </option>
                ))}
              </select>
            </div>

            {/* Count */}
            <div>
              <label className="block text-xs text-[#a1a1aa] mb-1 font-medium">
                Count
              </label>
              <input
                type="number"
                min={1}
                value={door.count}
                onChange={(e) =>
                  updateDoor(index, {
                    count: Math.max(1, parseInt(e.target.value) || 1),
                  })
                }
                className="w-full px-2.5 py-1.5 bg-[#09090b] border border-[#27272a] rounded-md text-sm text-white focus:outline-none focus:border-[#2563eb] transition-colors"
              />
            </div>

            {/* Stile */}
            <div>
              <label className="block text-xs text-[#a1a1aa] mb-1 font-medium">
                Stile
              </label>
              <select
                value={door.stile}
                onChange={(e) => updateDoor(index, { stile: e.target.value })}
                className="w-full px-2.5 py-1.5 bg-[#09090b] border border-[#27272a] rounded-md text-sm text-white focus:outline-none focus:border-[#2563eb] transition-colors"
              >
                {STILE_OPTIONS.map((stile) => (
                  <option key={stile} value={stile}>
                    {stile}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Hardware checkboxes */}
          <div>
            <label className="block text-xs text-[#a1a1aa] mb-2 font-medium">
              Hardware
            </label>
            <div className="grid grid-cols-2 gap-x-4 gap-y-2">
              {HARDWARE_OPTIONS.map((hw) => (
                <label
                  key={hw}
                  className="flex items-center gap-2 cursor-pointer group"
                >
                  <input
                    type="checkbox"
                    checked={!!door.hardware[hw]}
                    onChange={() => toggleHardware(index, hw)}
                    className="w-3.5 h-3.5 rounded border-[#3f3f46] bg-[#09090b] text-[#2563eb] focus:ring-[#2563eb] focus:ring-offset-0 cursor-pointer"
                  />
                  <span className="text-xs text-[#a1a1aa] group-hover:text-white transition-colors">
                    {hw}
                  </span>
                </label>
              ))}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
