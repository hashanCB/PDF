import { ArrowUpTrayIcon, XMarkIcon } from '@heroicons/react/24/solid';

interface FileUploadProps {
  onFileUpload: (files: FileList) => void;
  onClear: () => void;
}

export default function FileUpload({ onFileUpload, onClear }: FileUploadProps) {
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      onFileUpload(files);
      e.target.value = ''; // Reset input
    }
  };

  return (
    <div className="flex flex-col sm:flex-row gap-4 items-center">
      <label className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg cursor-pointer hover:bg-purple-700 transition-colors">
        <ArrowUpTrayIcon className="w-5 h-5" />
        <span>Upload Excel Files</span>
        <input
          type="file"
          multiple
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          className="hidden"
        />
      </label>

      <button
        onClick={onClear}
        className="flex items-center gap-2 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition-colors"
      >
        <XMarkIcon className="w-5 h-5" />
        <span>Clear Data</span>
      </button>
    </div>
  );
} 