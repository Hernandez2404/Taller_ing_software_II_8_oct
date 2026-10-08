<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use App\Http\Requests\BulkStoreUsersRequest;
use App\Models\User;
use Carbon\Carbon;
use Illuminate\Http\JsonResponse;
use Illuminate\Support\Facades\Hash;
use Illuminate\Http\Request;


class UserController extends Controller
{
    /**
     * Obtener cantidad de elementos por página con límite de seguridad (máximo 200).
     */
    private function getPerPage(Request $request): int
    {
        $perPage = (int) $request->query('per_page', 50);
        return max(1, min($perPage, 200));
    }

    public function index(Request $request): JsonResponse
    {
        $perPage = $this->getPerPage($request);
        $users = User::paginate($perPage);

        return response()->json($users);
    }

    public function emails(Request $request): JsonResponse
    {
        $perPage = $this->getPerPage($request);
        $users = User::select(['id', 'email'])->paginate($perPage);

        return response()->json($users);
    }

    public function overTwenty(Request $request): JsonResponse
    {
        $perPage = $this->getPerPage($request);
        $cutoff = Carbon::now()->subYears(20)->startOfDay();

        $users = User::where('birth_date', '<=', $cutoff)
            ->paginate($perPage);

        return response()->json($users);
    }

    public function bulkStore(BulkStoreUsersRequest $request): JsonResponse
    {
        $created = [];

        foreach ($request->validated()['users'] as $userData) {
            $created[] = User::create([
                'name' => $userData['name'],
                'email' => $userData['email'],
                'birth_date' => $userData['birth_date'],
                'password' => Hash::make($userData['password'] ?? 'password'),
            ]);
        }

        return response()->json([
            'message' => 'Se crearon 3 usuarios correctamente.',
            'users' => $created,
        ], 201);
    }
}
