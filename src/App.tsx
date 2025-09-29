import React, { useEffect, useState } from "react";

interface Endpoint {
  ADevice: string;
  APort: string;
  ZDevice: string;
  ZPort: string;
}

function App() {
  const [data, setData] = useState<Endpoint[]>([]);

  useEffect(() => {
    // Load sample JSON from public folder
    fetch("sample.json")
      .then((res) => res.json())
      .then(setData)
      .catch((err) => console.error("Failed to load JSON", err));
  }, []);

  return (
    <div className="min-h-screen bg-gray-900 text-white p-6">
      <h1 className="text-3xl font-bold mb-6 text-blue-400">
        OLS Optical Link Summary
      </h1>

      <div className="overflow-x-auto bg-gray-800 rounded-xl shadow-lg p-4">
        <table className="table-auto w-full text-sm">
          <thead className="bg-gray-700 text-blue-300">
            <tr>
              <th className="px-2 py-1">#</th>
              <th className="px-2 py-1">ADevice</th>
              <th className="px-2 py-1">APort</th>
              <th className="px-2 py-1">ZDevice</th>
              <th className="px-2 py-1">ZPort</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, i) => (
              <tr key={i} className="odd:bg-gray-800 even:bg-gray-700">
                <td className="px-2 py-1">{i + 1}</td>
                <td className="px-2 py-1">{row.ADevice}</td>
                <td className="px-2 py-1">{row.APort}</td>
                <td className="px-2 py-1">{row.ZDevice}</td>
                <td className="px-2 py-1">{row.ZPort}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="mt-6 flex gap-2">
        <button className="bg-blue-600 px-3 py-1 rounded">Copy All</button>
        <button className="bg-green-600 px-3 py-1 rounded">Export CSV</button>
        <button className="bg-gray-600 px-3 py-1 rounded">Refresh</button>
      </div>
    </div>
  );
}

export default App;
