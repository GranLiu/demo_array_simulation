# Compact Array Simulation

This is a matlab demo for creating a compact array with rectangular lattice, composed of half-wavelength dipoles, and exporting data files.

*Supported by Wireless Network RAN Research Department, Huawei Technologies CO., Ltd, Shanghai, China; School of Information and Communications Engineering, Xi'an Jiaotong University.*

**Athour: Youngxi Liu 刘雍熙  E-mail: <liu_xii@foxmail.com>**

***

## Change Log

### v0.1

- initialize project and add description

## File Structure
```bash
# dir for EM simulation #######################
./array_simulation
# CST array simulation project ################
├── demo_array.cst
│
# matlab file used to generate project ########
└── demo_array.m     

# dir for data export & process ###############
./post_processing
# export farfield data to txt #################
├── demo_export_farfield.m
│
# export 1D efficiency data to txt ############
└── demo_export_eff.m

```

## Work Flow
1. Open CST software,
2. Open ./array_simulation/demo_array.m, change n_x (n_y),
3. Run matlab script, an array composed of dipoles will be automatically generated and simulated,
4. Open ./post_processing/demo_export_farfield.m (demo_export_eff.m), change n_x (n_y),
5. Run script, data files will be exported into ./post_processing/data_***
- More information can be found in the notes in scripts.

## Q&A
- to be filled