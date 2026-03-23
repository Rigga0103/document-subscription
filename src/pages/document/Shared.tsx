import { useState, useEffect } from "react";
import {
  Search,
  FileText,
  Mail,
  MessageCircle,
  Loader,
  RefreshCw,
} from "lucide-react";
import useHeaderStore from "../../store/headerStore";
import { formatDate } from "../../utils/dateFormatter";

interface ShareHistoryItem {
  id: string;
  shareNo: string;
  dateTime: string;
  docSerial: string;
  docName: string;
  docFile: string;
  sharedVia: string;
  recipientName: string;
  contactInfo: string;
  email?: string;
  documentType?: string;
  category?: string;
  serialNo?: string;
  image?: string;
  sourceSheet?: string;
  shareMethod?: string;
  number?: string;
}

const SharedDocuments = () => {
  const { setTitle } = useHeaderStore();
  const [shareHistory, setShareHistory] = useState<ShareHistoryItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState("");
  const [error, setError] = useState<string | null>(null);

  const GOOGLE_SCRIPT_URL = import.meta.env.VITE_GOOGLE_SCRIPT_URL || "";

  useEffect(() => {
    setTitle("Share History");
    fetchShareHistoryFromGoogleSheets();
  }, [setTitle]);

  const fetchShareHistoryFromGoogleSheets = async () => {
    try {
      setLoading(true);
      setError(null);

      if (!GOOGLE_SCRIPT_URL) {
        throw new Error("Google Script URL is not defined");
      }

      const url = `${GOOGLE_SCRIPT_URL}?sheet=Shared Documents&_t=${Date.now()}`;

      const res = await fetch(url);

      if (!res.ok) {
        throw new Error(`Error: ${res.status}`);
      }

      const json = await res.json();

      if (!json?.data || !Array.isArray(json.data)) {
        throw new Error("Invalid data from Shared Documents sheet");
      }

      // Fetch ALL rows from "Shared Documents" sheet
      const rows = json.data;

      const transformedData = rows.slice(1).map((row: any[], index: number) => {
        return {
          id: `share-${index}`,
          shareNo: `SH-${String(index + 1).padStart(3, "0")}`,
          dateTime: row[0] || "",           // Timestamp
          email: row[1] || "",               // Email
          recipientName: row[2] || "",       // Name
          docName: row[3] || "",             // Document Name
          documentType: row[4] || "",        // Document Type
          category: row[5] || "",            // Category
          docSerial: row[6] || "",           // Serial No
          image: row[7] || "",               // Image
          sourceSheet: row[8] || "Shared Documents",  // Source Sheet
          sharedVia: row[9] || "",           // Share Method
          contactInfo: row[10] || row[1] || "N/A",    // Number (or Email fallback)
        };
      });

      setShareHistory(transformedData);
    } catch (err: any) {
      setError(err.message || "Failed to fetch data");
    } finally {
      setLoading(false);
    }
  };

  const filteredData = shareHistory.filter(
    (item) =>
      (item.shareNo || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.docName || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.recipientName || "")
        .toLowerCase()
        .includes(searchTerm.toLowerCase()) ||
      (item.docSerial || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.contactInfo || "").toLowerCase().includes(searchTerm.toLowerCase()),
  );

  const handleRefresh = () => {
    fetchShareHistoryFromGoogleSheets();
  };

  if (loading) {
    return (
      <div className="flex justify-center items-center min-h-screen">
        <div className="flex flex-col items-center gap-4">
          <Loader className="h-8 w-8 text-indigo-600 animate-spin" />
          <p className="text-gray-600">
            Loading share history from Google Sheets...
          </p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen">
        <div className="bg-red-50 border border-red-200 rounded-lg p-6 max-w-md">
          <div className="text-red-600 font-medium mb-2">
            Error Loading Data
          </div>
          <p className="text-gray-700 mb-4">{error}</p>
          <button
            onClick={handleRefresh}
            className="flex items-center justify-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors"
          >
            <RefreshCw size={16} />
            Try Again
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="space-y-6 pb-20">
      {/* Header & Search */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-white p-4 rounded-xl shadow-sm border border-gray-100">
        <div>
          <h1 className="text-2xl font-bold text-gray-800">Share History</h1>
          <p className="text-sm text-gray-500 mt-1">
            Track all shared documents from Google Sheets
          </p>
          {shareHistory.length > 0 && (
            <p className="text-xs text-gray-400 mt-1">
              {shareHistory.length} record{shareHistory.length !== 1 ? "s" : ""}{" "}
              found
            </p>
          )}
        </div>
        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-5 w-5" />
            <input
              type="text"
              placeholder="Search by Name, Document, Serial, Email..."
              className="pl-10 pr-4 py-2.5 w-full border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-gray-50"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          <button
            onClick={handleRefresh}
            className="flex items-center gap-2 px-4 py-2.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors shadow-sm hover:shadow"
          >
            <RefreshCw size={16} className={loading ? "animate-spin" : ""} />
            Refresh
          </button>
        </div>
      </div>

      {/* Desktop Table */}
      <div className="hidden md:block bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto max-h-[calc(100vh-250px)]">
          <table className="w-full text-left border-collapse">
            <thead className="bg-gray-50 border-b border-gray-100 sticky top-0 z-10">
              <tr>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  #
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Date
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Email
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Name
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Document Name
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Document Type
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Category
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Serial No.
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Share Method
                </th>
                <th className="p-3 text-xs font-semibold text-gray-500 uppercase tracking-wider whitespace-nowrap">
                  Number
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {filteredData.map((item) => (
                <tr
                  key={item.id}
                  className="hover:bg-indigo-50/10 transition-colors"
                >
                  <td className="p-3 text-sm font-medium text-indigo-600">
                    {item.shareNo}
                  </td>
                  <td className="p-3 text-sm text-gray-600 whitespace-nowrap">
                    {item.dateTime ? formatDate(item.dateTime) : "N/A"}
                  </td>
                  <td className="p-3 text-sm text-gray-600">
                    {item.email || "N/A"}
                  </td>
                  <td className="p-3 text-sm text-gray-900 font-medium">
                    {item.recipientName || "N/A"}
                  </td>
                  <td className="p-3 text-sm text-gray-900">
                    <div className="flex items-center gap-2">
                      <FileText size={16} className="text-gray-400" />
                      <span
                        className="truncate max-w-[180px]"
                        title={item.docName}
                      >
                        {item.docName}
                      </span>
                    </div>
                  </td>
                  <td className="p-3 text-sm text-gray-600">
                    {item.documentType || "N/A"}
                  </td>
                  <td className="p-3 text-sm text-gray-600">
                    {item.category || "N/A"}
                  </td>
                  <td className="p-3 text-sm text-gray-900 font-mono">
                    {item.docSerial || "N/A"}
                  </td>
                  <td className="p-3 text-sm text-gray-600">
                    <div className="flex items-center gap-2">
                      {item.sharedVia === "Email" ? (
                        <Mail size={16} className="text-blue-500" />
                      ) : item.sharedVia === "WhatsApp" ? (
                        <MessageCircle size={16} className="text-green-500" />
                      ) : null}
                      {item.sharedVia}
                    </div>
                  </td>
                  <td className="p-3 text-sm text-gray-600">
                    {item.contactInfo || "N/A"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {filteredData.length === 0 && searchTerm ? (
          <div className="p-12 text-center text-gray-500">
            <Search className="mx-auto h-12 w-12 text-gray-300 mb-3" />
            <p>No share history found matching "{searchTerm}"</p>
            <p className="text-sm text-gray-400 mt-1">
              Try a different search term
            </p>
          </div>
        ) : filteredData.length === 0 ? (
          <div className="p-12 text-center text-gray-500">
            <FileText className="mx-auto h-12 w-12 text-gray-300 mb-3" />
            <p>No share history found in Google Sheets</p>
            <p className="text-sm text-gray-400 mt-1">
              The "Shared Documents" sheet might be empty
            </p>
          </div>
        ) : null}
      </div>

      {/* Mobile View */}
      <div className="md:hidden space-y-4">
        {filteredData.map((item) => (
          <div
            key={item.id}
            className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex flex-col gap-3"
          >
            <div className="flex justify-between items-start">
              <span className="text-xs font-bold text-indigo-600 bg-indigo-50 px-2 py-1 rounded">
                {item.shareNo}
              </span>
              <p className="text-xs text-gray-400">
                {item.dateTime ? formatDate(item.dateTime) : "N/A"}
              </p>
            </div>

            <div className="border-t border-gray-50 pt-3 flex flex-col space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Email:</span>
                <span className="text-sm text-gray-900 text-right">
                  {item.email || "N/A"}
                </span>
              </div>

              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Name:</span>
                <span className="text-sm text-gray-900 font-medium text-right">
                  {item.recipientName || "N/A"}
                </span>
              </div>

              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Document:</span>
                <div className="text-right">
                  <span className="text-sm font-medium text-gray-900 block">
                    {item.docName}
                  </span>
                  {item.documentType && (
                    <span className="text-xs text-gray-400 block">
                      {item.documentType}
                    </span>
                  )}
                </div>
              </div>

              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Category:</span>
                <span className="text-sm text-gray-600 text-right">
                  {item.category || "N/A"}
                </span>
              </div>

              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Serial No:</span>
                <span className="text-sm text-gray-900 font-mono text-right">
                  {item.docSerial || "N/A"}
                </span>
              </div>

              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Share Method:</span>
                <div className="flex items-center gap-1.5 text-sm text-gray-700">
                  {item.sharedVia === "Email" ? (
                    <Mail size={14} className="text-blue-500" />
                  ) : item.sharedVia === "WhatsApp" ? (
                    <MessageCircle size={14} className="text-green-500" />
                  ) : null}
                  {item.sharedVia}
                </div>
              </div>

              <div className="flex items-center justify-between">
                <span className="text-xs text-gray-500">Number:</span>
                <span className="text-sm text-gray-600 text-right">
                  {item.contactInfo || "N/A"}
                </span>
              </div>
            </div>
          </div>
        ))}

        {filteredData.length === 0 && searchTerm ? (
          <div className="p-8 text-center text-gray-500 bg-white rounded-xl border border-gray-100">
            <Search className="mx-auto h-12 w-12 text-gray-300 mb-3" />
            <p>No records found for "{searchTerm}"</p>
            <p className="text-sm text-gray-400 mt-1">
              Try a different search term
            </p>
          </div>
        ) : filteredData.length === 0 ? (
          <div className="p-8 text-center text-gray-500 bg-white rounded-xl border border-gray-100">
            <FileText className="mx-auto h-12 w-12 text-gray-300 mb-3" />
            <p>No records found</p>
            <p className="text-sm text-gray-400 mt-1">
              Data source: Google Sheets - "Shared Documents" sheet
            </p>
          </div>
        ) : null}
      </div>
    </div>
  );
};
export default SharedDocuments;
